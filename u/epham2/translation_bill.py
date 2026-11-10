# requirements:
# stripe
# requests
# pycountry
# wmill

"""
Raise a translation invoice when an application is flagged for billing.

FLOW
----
    Someone sets the application status to "Bill_for_translation"
      -> automation calls this webhook with the application id
      -> read pages from custom field 153837 (translation_pages_required)
      -> Stripe invoice for pages x $110 CAD, emailed to the applicant
      -> application stays on Bill_for_translation until payment

    Payment -> 20_payment_received.py restores the status from custom field
               153838 (prior_status) and emails the translation coordinator.

    Everything after that is manual.

DEPENDENCY WORTH KNOWING
------------------------
Custom field 153838 (prior_status) must be populated by whoever sets the
Bill_for_translation status. This script cannot capture it — by the time the
webhook fires, the status has already changed, so the previous value is gone.

If 153838 is empty the invoice is still raised (refusing would make the
reviewer's action silently do nothing), but the run reports it loudly and
payment will have nothing to restore to. Worth wiring into whatever sets the
billing status.

requirements:
stripe
requests
pycountry
wmill
"""

import json
import re
from datetime import datetime, timezone

import requests
import stripe
import wmill

# ======================================================================
# CONFIG
# ======================================================================

BASE = "https://nnas.smapply.io/api"
TIMEOUT = 60

TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# Restricted Stripe key (rk_...), NOT the key SMApply uses for application fees.
# Separate keys keep the two flows distinguishable in Stripe's request logs and
# independently rotatable. Needs write on Customers / Invoices / Invoice Items.
STRIPE_VAR = "u/epham2/STRIPE_SECRET_KEY"

RATE_CENTS = 12_500  # $125.00 CAD per page
MAX_PAGES = 15
CURRENCY = "cad"
DAYS_UNTIL_DUE = 30

# "reject" fails loudly over the cap; "clamp" bills 15 and absorbs the rest.
OVER_LIMIT_BEHAVIOUR = "reject"

# --- Tax -----------------------------------------------------------------
#
#   "none"       no tax lines at all
#   "rate"       a fixed Stripe Tax Rate applied to the line item. Simple,
#                works today, needs no address.
#   "automatic"  Stripe Tax computes it. Needs THREE things beyond this
#                setting, and finalization fails without them:
#                  1. Stripe Tax activated (Settings -> Tax) — it is a paid
#                     add-on, not on by default
#                  2. at least one tax registration configured, so Stripe knows
#                     where you are obliged to collect
#                  3. a customer address Stripe can resolve. Ours come from
#                     SMApply, which is why ADDRESS_TASK_ID exists below.
#
# Which is correct is a tax question, not a technical one. Place-of-supply
# rules for services delivered to applicants outside Canada are not something
# to guess at — settle it with finance or your accountant first.
TAX_MODE = "automatic"

TAX_RATE_ID = None  # "txr_..." — required when TAX_MODE = "rate"

# Whether the $110 already includes tax, or tax is added on top. Getting this
# backwards is a pricing error, not a display one.
TAX_BEHAVIOR = "exclusive"  # "exclusive" | "inclusive"

# --- Customer address, for TAX_MODE = "automatic" ------------------------
# The applicant's address is not on the application object — it lives in the
# Personal Information task. Leave ADDRESS_TASK_ID at 0 to skip address sync.
ADDRESS_TASK_ID = 0  # e.g. 2044268 "Personal Information"
ADDRESS_FIELD_KEYS = {
    "line1": "odQBnoDDlJ_0",
    "line2": "odQBnoDDlJ_1",
    "city": "odQBnoDDlJ_2",
    "state": "odQBnoDDlJ_3",
    "postal_code": "odQBnoDDlJ_4",
    "country": "odQBnoDDlJ_5",
}

# The webhook template renders the country dropdown as a NAME ("Benin"), while
# the API returns its option index (18). Passing the address in the automation
# payload therefore avoids maintaining an index table — see 00_SETUP.md.
# Names are converted to ISO-3166 alpha-2 by to_iso2() below.
DEFAULT_COUNTRY = ""  # blank = refuse rather than guess a country

# Names that pycountry does not resolve on its own, or that SMApply spells
# differently. Extend as you meet them — a wrong country is a wrong tax rate.
COUNTRY_NAME_OVERRIDES = {
    "united states": "US",
    "usa": "US",
    "u.s.a.": "US",
    "united kingdom": "GB",
    "uk": "GB",
    "great britain": "GB",
    "south korea": "KR",
    "korea, south": "KR",
    "republic of korea": "KR",
    "north korea": "KP",
    "russia": "RU",
    "vietnam": "VN",
    "viet nam": "VN",
    "laos": "LA",
    "iran": "IR",
    "syria": "SY",
    "bolivia": "BO",
    "venezuela": "VE",
    "tanzania": "TZ",
    "moldova": "MD",
    "taiwan": "TW",
    "ivory coast": "CI",
    "cote d'ivoire": "CI",
    "côte d'ivoire": "CI",
    "cape verde": "CV",
    "cabo verde": "CV",
    "democratic republic of the congo": "CD",
    "drc": "CD",
    "congo, democratic republic of the": "CD",
    "republic of the congo": "CG",
    "congo": "CG",
    "macedonia": "MK",
    "north macedonia": "MK",
    "czech republic": "CZ",
    "czechia": "CZ",
    "burma": "MM",
    "myanmar": "MM",
    "palestine": "PS",
    "swaziland": "SZ",
    "eswatini": "SZ",
    "turkey": "TR",
    "türkiye": "TR",
    "turkiye": "TR",
    "cape verde islands": "CV",
    "east timor": "TL",
    "timor-leste": "TL",
    "philippines": "PH",
    "the philippines": "PH",
    "netherlands": "NL",
    "the netherlands": "NL",
    "holland": "NL",
    "bahamas": "BS",
    "the bahamas": "BS",
    "gambia": "GM",
    "the gambia": "GM",
}

# Stripe product tax code for the service being sold. "txcd_20060045" is
# Stripe's general services code; there are more specific ones. Leave blank to
# use the account default. Check the current list at
# https://docs.stripe.com/tax/tax-codes — the right code is a finance decision.
PRODUCT_TAX_CODE = "txcd_10000000"

# --- SMApply -------------------------------------------------------------
PAGES_FIELD = 153837  # translation_pages_required (text holding a number)
PRIOR_STATUS_FIELD = 153838  # prior_status (status id to restore on payment)

# The status that means "bill this applicant". Look it up with
# GET /api/stages/<stage_id>/statuses/ and set it here.
BILL_STATUS_ID = 366184  # "Bill_for_translation"

# Never write these — 152033 "Payment date" is the APPLICATION FEE date, owned
# by the existing date-stamping workflow. Same for the other two.
RESERVED_CUSTOM_FIELDS = {152033, 153648, 153677}

INVOICE_FOOTER = (
    "This charge covers certified translation of documents submitted with your "
    "NNAS application. Your application is on hold until payment is received. "
    "Do not reply to this email — contact NNAS through your applicant portal."
)
CF_LIMIT = 30  # Stripe caps invoice custom_field name and value at 30 chars


# ======================================================================
# SMApply
# ======================================================================


def _sm_headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def _body(resp: requests.Response):
    """SMApply 500s return HTML. Never assume JSON."""
    try:
        return resp.json()
    except Exception:
        return {"text": (resp.text or "")[:800]}


def get_application(token: str, application_id: int) -> dict:
    r = requests.get(
        f"{BASE}/applications/{application_id}/",
        headers=_sm_headers(token),
        timeout=TIMEOUT,
    )
    if not (200 <= r.status_code < 300):
        raise RuntimeError(
            f"GET application {application_id} failed: "
            f"HTTP {r.status_code} "
            f"{json.dumps(_body(r), default=str)[:500]}"
        )
    return _body(r)


def custom_fields(app: dict) -> dict:
    """{field_id: value} for everything on the application."""
    return {
        f.get("id"): str(f.get("value") or "")
        for f in (app.get("custom_fields") or [])
        if f.get("id")
    }


def current_status(app: dict):
    st = app.get("status")
    if isinstance(st, dict):
        return st.get("id")
    return st if isinstance(st, int) else None


def extract_applicant(app: dict, application_id: int) -> dict:
    owner = app.get("applicant") or app.get("owner") or app.get("user") or {}
    if not isinstance(owner, dict):
        owner = {}
    email = owner.get("email") or app.get("email")
    name = (
        f"{owner.get('first_name', '')} {owner.get('last_name', '')}"
    ).strip() or owner.get("full_name")
    if not email:
        raise RuntimeError(
            f"No applicant email on application {application_id}. Cannot invoice. "
            f"Keys present: {sorted(app.keys())}"
        )
    return {
        "email": email,
        "name": name or "(name not on file)",
        "reference_id": app.get("reference_id") or str(application_id),
    }


def parse_pages(raw: str, application_id: int) -> int:
    """
    153837 is a free-text field, so it may hold "5", "5 pages", " 5 " or junk.
    Take the first integer; refuse anything with no digits rather than guessing.
    """
    m = re.search(r"\d+", str(raw or ""))
    if not m:
        raise ValueError(
            f"Custom field {PAGES_FIELD} on application {application_id} is "
            f"{raw!r} — no page count found. Nothing invoiced."
        )
    pages = int(m.group())
    if pages < 1:
        raise ValueError(
            f"Page count on application {application_id} is {pages}. Nothing invoiced."
        )
    if pages > MAX_PAGES:
        if OVER_LIMIT_BEHAVIOUR == "clamp":
            print(f"NOTE: {pages} pages clamped to the {MAX_PAGES}-page cap")
            return MAX_PAGES
        raise ValueError(
            f"{pages} pages exceeds the {MAX_PAGES}-page cap "
            f"(${MAX_PAGES * RATE_CENTS / 100:,.2f}). Quote this one manually."
        )
    return pages


# ======================================================================
# Stripe
# ======================================================================


def _meta(obj) -> dict:
    """
    Metadata off a Stripe object, as a real dict.

    stripe-python 15 returns a StripeObject here, which looks dict-like but is
    not — `.get()` raises AttributeError rather than returning None. Attribute
    access on a missing metadata key would raise too, so convert once and use
    plain dict semantics everywhere after.
    """
    m = getattr(obj, "metadata", None)
    if not m:
        return {}
    if isinstance(m, dict):
        return dict(m)
    to_dict = getattr(m, "to_dict", None)
    if callable(to_dict):
        return to_dict()
    try:
        return dict(m)
    except Exception:
        return {}


def load_blob(additional_content) -> dict:
    """
    Parse SMApply's `additional_content` into a dict.

    It arrives as a JSON *string*, and it is frequently not valid JSON. The
    payload from application 35367543 contained:

        { "id" : NNAS-T-0000000032, "address_line1": "20158 82 ave", ... } }

    — an unquoted value and an unbalanced brace, both from the automation
    template. `json.loads` refuses it outright.

    So: try strict JSON first, and fall back to pulling out every quoted
    "key": "value" pair. That fallback happens to skip exactly the malformed
    entry (the unquoted id) and recover every field we actually want. Fix the
    template anyway — see 00_SETUP.md — but do not depend on it being fixed.
    """
    if not additional_content:
        return {}
    if isinstance(additional_content, dict):
        return additional_content

    text = str(additional_content).strip()
    try:
        blob = json.loads(text)
        if isinstance(blob, str):  # sometimes double-encoded
            blob = json.loads(blob)
        if isinstance(blob, dict):
            return blob
    except Exception:
        pass

    salvaged = dict(re.findall(r'"(\w+)"\s*:\s*"([^"]*)"', text))
    if salvaged:
        print(
            f"NOTE: additional_content is not valid JSON; salvaged "
            f"{sorted(salvaged)} by pattern match. Fix the automation "
            f"template — quote every value and check the braces."
        )
    else:
        print(f"WARN: could not parse additional_content: {text[:200]!r}")
    return salvaged


def to_iso2(country) -> str:
    """
    Country name -> ISO-3166 alpha-2, which is what Stripe wants.

    The webhook posts the rendered label ("Benin"), so this is a name lookup,
    not an index lookup. Overrides are checked first because pycountry's
    official names differ from everyday usage for a fair few countries, and a
    country that silently resolves wrong is a tax rate that silently computes
    wrong.

    Returns "" rather than guessing when it cannot resolve.
    """
    if not country:
        return ""
    raw = str(country).strip()
    if len(raw) == 2 and raw.isalpha():
        return raw.upper()

    key = raw.lower().strip(" .")
    if key in COUNTRY_NAME_OVERRIDES:
        return COUNTRY_NAME_OVERRIDES[key]

    try:
        import pycountry

        hit = (
            pycountry.countries.get(name=raw)
            or pycountry.countries.get(common_name=raw)
            or pycountry.countries.get(official_name=raw)
        )
        if hit:
            return hit.alpha_2
        matches = pycountry.countries.search_fuzzy(raw)
        if matches:
            print(
                f"NOTE: country {raw!r} matched fuzzily to "
                f"{matches[0].name} ({matches[0].alpha_2}). Add it to "
                f"COUNTRY_NAME_OVERRIDES to make that explicit."
            )
            return matches[0].alpha_2
    except Exception as exc:
        print(f"WARN: country lookup for {raw!r} failed ({exc})")
    return ""


def build_address(
    line1="", line2="", city="", state="", postal_code="", country=""
) -> dict:
    """
    Shape an address for Stripe, or return {} if it is not usable.

    A partial or unresolvable address is discarded rather than sent. Stripe
    would accept it and compute tax for the wrong jurisdiction, which is worse
    than not computing at all.
    """
    iso = to_iso2(country) or DEFAULT_COUNTRY
    address = {
        "line1": str(line1 or "").strip(),
        "line2": str(line2 or "").strip() or None,
        "city": str(city or "").strip(),
        "state": str(state or "").strip() or None,
        "postal_code": str(postal_code or "").strip(),
        "country": iso,
    }

    missing = [k for k in ("line1", "city", "country") if not address[k]]
    if missing:
        print(
            f"WARN: address unusable, missing {missing} "
            f"(country input was {country!r}). Not syncing."
        )
        return {}
    return {k: v for k, v in address.items() if v}


def read_address(token: str, application_id: int) -> dict:
    """
    Fallback: read the address from the Personal Information task.

    Only used when the webhook did not carry one. Note the API returns the
    country as an option INDEX rather than a name, so this path needs
    ADDRESS_TASK_ID and will usually fail to resolve a country — prefer passing
    the address in the automation payload, where SMApply renders the label.
    """
    if not ADDRESS_TASK_ID:
        return {}

    r = requests.get(
        f"{BASE}/applications/{application_id}/tasks/{ADDRESS_TASK_ID}/",
        headers=_sm_headers(token),
        timeout=TIMEOUT,
    )
    if not (200 <= r.status_code < 300):
        print(f"WARN: address task {ADDRESS_TASK_ID} unreadable: HTTP {r.status_code}")
        return {}

    data = (_body(r) or {}).get("data") or {}

    def val(key):
        v = data.get(key)
        if isinstance(v, dict):
            v = v.get("response")
        if isinstance(v, list) and len(v) == 1:
            v = v[0]
        return v

    return build_address(
        line1=val(ADDRESS_FIELD_KEYS["line1"]),
        line2=val(ADDRESS_FIELD_KEYS["line2"]),
        city=val(ADDRESS_FIELD_KEYS["city"]),
        state=val(ADDRESS_FIELD_KEYS["state"]),
        postal_code=val(ADDRESS_FIELD_KEYS["postal_code"]),
        country=val(ADDRESS_FIELD_KEYS["country"]),
    )


def sync_customer_address(customer_id: str, address: dict):
    """
    Stripe determines invoice tax from the CUSTOMER's address. Ours may have
    been created by SMApply's fee flow without one, so set it when we have it.

    This does write to a customer SMApply may own — unlike metadata, an address
    is factual rather than system-specific, and an invoice that cannot compute
    tax is a harder failure than a shared address field.
    """
    if not address:
        if TAX_MODE == "automatic":
            raise RuntimeError(
                "TAX_MODE is 'automatic' but no usable address is available. "
                "Stripe cannot determine the tax jurisdiction, so the invoice "
                "would either fail to finalize or charge the wrong amount. "
                "Pass the address in the automation payload — see 00_SETUP.md."
            )
        return None
    try:
        stripe.Customer.modify(customer_id, address=address)
        print(
            f"Customer {customer_id} address set: {address.get('city')}, "
            f"{address.get('country')}"
        )
        return address
    except Exception as exc:
        if TAX_MODE == "automatic":
            # Better to stop than to issue an invoice taxed for the wrong place.
            raise RuntimeError(
                f"Could not set the customer address ({exc}). Automatic tax "
                f"cannot resolve without it. Check the address values — Stripe "
                f"validates postal codes against the country."
            )
        print(f"WARN: could not set customer address ({exc})")
        return None


def _cf(value: str) -> str:
    v = str(value or "")
    return v if len(v) <= CF_LIMIT else v[: CF_LIMIT - 1] + "\u2026"


def upsert_customer(applicant: dict, application_id: int) -> str:
    """
    Reuse the applicant's existing Stripe Customer if there is one.

    SMApply bills application fees through this same Stripe account, so a
    Customer for this person may already exist. Reusing it keeps one timeline
    per applicant instead of splitting them across two records.

    Customer.list, not Customer.search — list is strongly consistent, search is
    index-backed and lags.

    Deliberately does not write our metadata onto a customer we did not create;
    SMApply owns that object, and mutating another system's records is the kind
    of coupling that breaks during an unrelated upgrade months later.
    """
    email = applicant["email"]
    try:
        existing = stripe.Customer.list(email=email, limit=1)
        if existing.data:
            print(f"Reusing Stripe customer {existing.data[0].id} for {email}")
            return existing.data[0].id
    except Exception as exc:
        print(f"WARN: customer lookup by email failed ({exc}); creating new")

    return stripe.Customer.create(
        email=email,
        name=applicant["name"],
        metadata={
            "nnas_application_id": str(application_id),
            "nnas_reference_id": applicant["reference_id"],
        },
        idempotency_key=f"nnas-cust-{application_id}",
    ).id


def find_open_invoice(customer_id: str, application_id: int):
    """
    Is there already an UNPAID translation invoice for this applicant?

    Only `open` blocks. A previously PAID invoice does not — a second round of
    translation is a legitimate thing to bill for.

    EMPTY DRAFTS ARE DELETED, NOT TREATED AS BLOCKING. This script creates the
    invoice first and attaches the line item second, so any failure in between
    leaves a draft with no lines. If those counted as outstanding, one failed
    run would block the applicant from ever being billed, and the only symptom
    would be a permanent "already outstanding" that looks like correct
    behaviour. A draft carrying lines is a different matter and does block.

    Invoice.list is strongly consistent, unlike Invoice.search, so there is no
    window where a double-fire slips through.
    """
    try:
        invoices = stripe.Invoice.list(customer=customer_id, limit=100)
    except Exception as exc:
        print(
            f"WARN: could not list invoices for {customer_id} ({exc}). "
            f"Proceeding without duplicate detection."
        )
        return None

    for inv in invoices.auto_paging_iter():
        meta = _meta(inv)
        if meta.get("nnas_service") != "translation":
            continue
        if str(meta.get("nnas_application_id")) != str(application_id):
            continue

        if inv.status == "draft" and not (inv.amount_due or 0):
            print(f"Discarding empty draft {inv.id} left by an earlier failed run")
            try:
                stripe.Invoice.delete(inv.id)
            except Exception as exc:
                print(f"WARN: could not delete empty draft {inv.id}: {exc}")
            continue

        if inv.status in ("draft", "open"):
            return inv
    return None


def _create_invoice_item(
    customer_id: str,
    invoice_id: str,
    application_id: int,
    pages: int,
    amount_cents: int,
    idem: str,
):
    """
    Attach the charge to the draft invoice.

    Newer Stripe API versions removed the flat `unit_amount` parameter from
    InvoiceItem — this account is on 2026-07-29.dahlia, which rejects it with
    "unknown parameter: unit_amount. Did you mean unit_amount_decimal?".

    Preferred path keeps quantity x unit pricing so the applicant sees
    "5 x $110.00" on the invoice rather than one lump sum. If that shape is
    rejected too, fall back to a flat `amount`, which every version accepts.
    The description carries the breakdown either way, so the fallback is
    cosmetically poorer but never wrong.

    NOTE ON IDEMPOTENCY KEYS: `idem` is the invoice id, which is unique by
    construction, so a retry against the same invoice is idempotent while a new
    invoice always gets fresh keys. Each pricing shape also uses its OWN suffix,
    because Stripe rejects a key reused with a different request body — sharing
    one between the two attempts would turn a recoverable fallback into a hard
    error.
    """
    desc = (
        f"Certified translation — {pages} page{'s' if pages != 1 else ''} "
        f"@ ${RATE_CENTS / 100:.2f} CAD per page "
        f"= ${amount_cents / 100:,.2f} CAD"
    )
    base = dict(
        customer=customer_id,
        invoice=invoice_id,
        currency=CURRENCY,
        description=desc,
        metadata={"nnas_application_id": str(application_id), "nnas_pages": str(pages)},
    )
    if TAX_MODE == "rate":
        if not TAX_RATE_ID:
            raise RuntimeError(
                'TAX_MODE = "rate" but TAX_RATE_ID is unset. Create a Tax Rate '
                "in Settings -> Tax rates and put its txr_... id in CONFIG."
            )
        base["tax_rates"] = [TAX_RATE_ID]
    elif TAX_MODE == "automatic":
        # Stripe Tax needs a product tax code to know how the service is
        # treated. Without one it falls back to the account default, which may
        # not match how translation services are classified.
        base["tax_behavior"] = TAX_BEHAVIOR
        if PRODUCT_TAX_CODE:
            base["tax_code"] = PRODUCT_TAX_CODE

    try:
        item = stripe.InvoiceItem.create(
            **base,
            quantity=pages,
            unit_amount_decimal=str(RATE_CENTS),
            idempotency_key=f"{idem}-item-uad",
        )
        print(f"Invoice item created with quantity x unit_amount_decimal")
        return item
    except Exception as exc:
        msg = str(exc).lower()
        if "unknown parameter" not in msg and "parameter_unknown" not in msg:
            raise
        print(
            f"NOTE: unit pricing rejected ({exc}). Falling back to a flat "
            f"amount — the line will show one total, and the breakdown stays "
            f"in the description."
        )

    return stripe.InvoiceItem.create(
        **base, amount=amount_cents, idempotency_key=f"{idem}-item-amt"
    )


def create_invoice(
    customer_id: str,
    application_id: int,
    applicant: dict,
    pages: int,
    amount_cents: int,
):
    """
    Draft first, attach the line item to that specific invoice, then finalize.
    Creating the InvoiceItem first risks it being swept onto an unrelated draft.

    If the line item fails, the draft is deleted rather than left behind. An
    orphaned draft is invisible in the Stripe dashboard's default view and would
    otherwise sit on the customer forever.
    """
    ref = applicant["reference_id"]

    inv = stripe.Invoice.create(
        customer=customer_id,
        collection_method="send_invoice",
        days_until_due=DAYS_UNTIL_DUE,
        currency=CURRENCY,
        pending_invoice_items_behavior="exclude",
        description=(
            f"Document translation — {applicant['name']} "
            f"({applicant['email']}) — NNAS {ref} / application "
            f"{application_id}"
        ),
        footer=INVOICE_FOOTER,
        custom_fields=[
            {"name": "NNAS Reference", "value": _cf(ref)},
            {"name": "Application ID", "value": _cf(str(application_id))},
            {"name": "Applicant", "value": _cf(applicant["name"])},
        ],
        metadata={
            "nnas_service": "translation",
            "nnas_application_id": str(application_id),
            "nnas_reference_id": ref,
            "nnas_applicant_name": applicant["name"],
            "nnas_applicant_email": applicant["email"],
            "nnas_pages": str(pages),
        },
        automatic_tax={"enabled": TAX_MODE == "automatic"},
        # No idempotency_key here, deliberately.
        #
        # Stripe caches a key's response for 24 hours. When find_open_invoice
        # deletes an orphaned draft, a re-run within that window would replay
        # the cached response and hand back the id of the invoice we just
        # deleted — which then fails at finalize, permanently, for a day.
        #
        # find_open_invoice is the real guard and Invoice.list is strongly
        # consistent, so the only gap is two webhooks firing in the same
        # instant. Status changes here are human-driven and serial, so that is
        # a far smaller risk than a wedged applicant nobody can bill.
    )

    try:
        _create_invoice_item(
            customer_id, inv.id, application_id, pages, amount_cents, inv.id
        )
    except Exception:
        try:
            stripe.Invoice.delete(inv.id)
            print(f"Deleted draft {inv.id} after the line item failed")
        except Exception as cleanup_exc:
            print(f"WARN: draft {inv.id} could not be deleted: {cleanup_exc}")
        raise

    inv = stripe.Invoice.finalize_invoice(inv.id)
    stripe.Invoice.send_invoice(inv.id)  # emails the applicant
    return stripe.Invoice.retrieve(inv.id)


# ======================================================================
# Entrypoint
# ======================================================================


def main(id: int = None, additional_content: str = "", dry_run: bool = False, **kwargs):
    """
    Triggered by an automation on the application status becoming
    Bill_for_translation.

    Only three inputs. SMApply sends a fixed envelope — title, round, status,
    reference_id, user, token, award, user_id, date_created, last_edited — and
    **kwargs absorbs all of it. Declaring those fields individually only
    clutters the Windmill run form; nothing here reads them.

        id                  the application id
        additional_content  JSON string carrying the applicant's address.
                            Also the manual-override path: paste a blob into
                            this box to run by hand.

    Page count and application status are NOT in the payload, deliberately.
    Both are read from the API at run time. Choices render as label text in
    SMApply templates, so matching on them breaks silently the moment anyone
    edits the wording. The address is the opposite case — label text ("Canada")
    is exactly what Stripe needs, and the API's index form is what would need a
    lookup table.
    """
    application_id = id or kwargs.get("APPLICATION_ID") or kwargs.get("application_id")
    if not application_id:
        return {
            "status": "error",
            "message": "No application id. Bind the webhook 'id' field.",
        }
    application_id = int(application_id)

    blob = load_blob(additional_content)
    address_line1 = blob.get("address_line1", "")
    address_line2 = blob.get("address_line2", "")
    city = blob.get("city", "")
    province = blob.get("province", "") or blob.get("state", "")
    postal_code = blob.get("postal_code", "")
    country = blob.get("country", "")

    token = wmill.get_variable(TOKEN_VAR)
    app = get_application(token, application_id)
    fields = custom_fields(app)
    status_now = current_status(app)

    # Guard: only bill an application actually flagged for it. A stray webhook
    # call, a replay, or a mis-wired automation should not charge anybody.
    if BILL_STATUS_ID and status_now != BILL_STATUS_ID:
        return {
            "application_id": application_id,
            "skipped": "not flagged for billing",
            "current_status": status_now,
            "expected_status": BILL_STATUS_ID,
        }

    applicant = extract_applicant(app, application_id)
    pages = parse_pages(fields.get(PAGES_FIELD), application_id)
    amount_cents = pages * RATE_CENTS

    prior_status = fields.get(PRIOR_STATUS_FIELD, "").strip()
    prior_ok = prior_status.isdigit()
    if not prior_ok:
        # Not fatal — but payment will have nothing to restore to, and someone
        # will have to fix the status by hand. Better to know now than then.
        print(
            f"WARNING: custom field {PRIOR_STATUS_FIELD} (prior_status) on "
            f"application {application_id} is {prior_status!r}. Payment will "
            f"not be able to restore the status automatically. Populate it "
            f"wherever Bill_for_translation gets set."
        )

    summary = {
        "application_id": application_id,
        "reference_id": applicant["reference_id"],
        "applicant_name": applicant["name"],
        "applicant_email": applicant["email"],
        "pages": pages,
        "rate_cad": RATE_CENTS / 100,
        "amount_cad": amount_cents / 100,
        "prior_status": prior_status or None,
        "prior_status_usable": prior_ok,
        "tax_mode": TAX_MODE,
        "address": build_address(
            address_line1, address_line2, city, province, postal_code, country
        )
        or None,
        "country_input": country or None,
        "dry_run": dry_run,
    }

    if dry_run:
        summary["note"] = "dry run — no Stripe or SMApply writes"
        print(json.dumps(summary, indent=2, default=str))
        return summary

    stripe.api_key = wmill.get_variable(STRIPE_VAR)
    customer_id = upsert_customer(applicant, application_id)

    if TAX_MODE == "automatic":
        address = build_address(
            address_line1, address_line2, city, province, postal_code, country
        )
        if not address:
            address = read_address(token, application_id)  # fallback
        sync_customer_address(customer_id, address)

    open_inv = find_open_invoice(customer_id, application_id)
    if open_inv:
        summary.update(
            {
                "skipped": "unpaid translation invoice already outstanding",
                "existing_invoice": open_inv.id,
                "existing_amount_cad": (open_inv.amount_due or 0) / 100,
                "hosted_invoice_url": open_inv.hosted_invoice_url,
                "note": (
                    "Void the existing invoice in Stripe before raising "
                    "another — for example if the page count changed."
                ),
            }
        )
        print(json.dumps(summary, indent=2, default=str))
        return summary

    inv = create_invoice(customer_id, application_id, applicant, pages, amount_cents)

    summary.update(
        {
            "stripe_customer_id": customer_id,
            "stripe_invoice_id": inv.id,
            "hosted_invoice_url": inv.hosted_invoice_url,
            "invoice_pdf_url": inv.invoice_pdf,
            "sent_to": applicant["email"],
            "tax_mode": TAX_MODE,
            "tax_cad": (getattr(inv, "tax", None) or 0) / 100,
            "total_cad": (getattr(inv, "total", None) or 0) / 100,
            "created_at": datetime.now(timezone.utc).isoformat(),
        }
    )
    total = (getattr(inv, "total", None) or amount_cents) / 100
    tax = (getattr(inv, "tax", None) or 0) / 100
    print(
        f"Invoice {inv.id}: subtotal ${amount_cents / 100:,.2f} + tax "
        f"${tax:,.2f} = ${total:,.2f} CAD, sent to {applicant['email']} — "
        f"{inv.hosted_invoice_url}"
    )
    if TAX_MODE == "automatic" and not tax:
        print(
            "NOTE: automatic tax computed $0.00. That may be correct (a "
            "non-taxable jurisdiction), or it may mean no registration "
            "covers this customer's location. Check Settings -> Tax."
        )
    print(json.dumps(summary, indent=2, default=str))
    return summary

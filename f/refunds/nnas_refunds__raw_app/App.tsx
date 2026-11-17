import React, { useEffect, useState } from 'react'
import { backend } from './wmill'
import './index.css'

// NNAS Refunds: one app for every role. The backend decides what each person
// may do (identity comes from the server-side ctx email, not from this page);
// this page only hides buttons that would be refused anyway.

type Roles = {
    email: string
    admin: boolean
    can_request: boolean
    can_approve: boolean
    can_process: boolean
    can_view_all: boolean
}

type Me = {
    you_are: string
    roles: Roles
    test_mode: boolean
    test_solo: boolean
    test_application_ids: number[]
}

type Row = Record<string, any>

const REASONS = [
    'Duplicate payment',
    'Overpayment',
    'Processing error',
    'Service not provided',
    'Withdrawal',
    'Other',
]

const STATUS_LABEL: Record<string, string> = {
    pending_approval: 'Pending approval',
    approved: 'Approved, awaiting refund',
    rejected: 'Rejected',
    refunded: 'Refunded, issue pending',
    closed: 'Closed',
}

// Runnable IDs as shown on the badges in the editor's BACKEND panel.
// If you delete and re-add a runnable it gets a new letter: change it here.
//
// These are called by ID on purpose. `backend` answers to ANY name, so a
// name that does not exist does not fail, it just never returns.
const RUNNABLE_ID = {
    whoami: 'a',
    list_requests: 'b',
    create_request: 'c',
    decide_request: 'd',
    complete_and_lock: 'e',
    find_payments: 'g',
}

const B = backend as any
const call = (id: string) => (args: Record<string, any> = {}) => B[id](args)

const api = {
    whoami: call(RUNNABLE_ID.whoami),
    list_requests: call(RUNNABLE_ID.list_requests),
    create_request: call(RUNNABLE_ID.create_request),
    decide_request: call(RUNNABLE_ID.decide_request),
    complete_and_lock: call(RUNNABLE_ID.complete_and_lock),
    find_payments: call(RUNNABLE_ID.find_payments),
}

function errText(e: any): string {
    const raw =
        e?.message ?? e?.error?.message ?? e?.result?.error?.message ??
        (typeof e === 'string' ? e : JSON.stringify(e))
    const lines = String(raw).trim().split('\n').map((l) => l.trim()).filter(Boolean)
    const last = lines[lines.length - 1] || 'Unknown error'
    return last.replace(/^(ValueError|PermissionError|RuntimeError|Exception):\s*/, '')
}

function when(s?: string | null): string {
    if (!s) return ''
    const d = new Date(s)
    return isNaN(d.getTime()) ? String(s) : d.toLocaleString('en-CA', { dateStyle: 'medium', timeStyle: 'short' })
}

function Badge({ status }: { status: string }) {
    return <span className={`badge badge-${status}`}>{STATUS_LABEL[status] ?? status}</span>
}

function Notice({ kind, text }: { kind: 'ok' | 'err' | 'warn'; text: string }) {
    if (!text) return null
    return <div className={`notice notice-${kind}`}>{text}</div>
}

// ====================================================================== App

export default function App() {
    const [me, setMe] = useState<Me | null>(null)
    const [loadErr, setLoadErr] = useState('')
    const [tab, setTab] = useState<'requests' | 'new'>('requests')
    const [refreshKey, setRefreshKey] = useState(0)

    useEffect(() => {
        let done = false
        const t = setTimeout(() => {
            if (!done) {
                setLoadErr(
                    `No answer from runnable "${RUNNABLE_ID.whoami}" after 60 seconds. Check that the ` +
                    'letters in RUNNABLE_ID (top of App.tsx) match the badges in the BACKEND panel.'
                )
            }
        }, 60000)
        ;(async () => {
            try {
                const r = await api.whoami({})
                done = true
                if (!r || typeof r !== 'object' || !r.roles) {
                    throw new Error(`whoami returned an unexpected result: ${JSON.stringify(r) ?? 'undefined'}`)
                }
                setMe(r)
            } catch (e) {
                done = true
                setLoadErr(errText(e))
            } finally {
                clearTimeout(t)
            }
        })()
        return () => clearTimeout(t)
    }, [])

    if (loadErr) return <div className="page"><Notice kind="err" text={`Could not load: ${loadErr}`} /></div>
    if (!me) return <div className="page muted">Loading...</div>

    const r = me.roles
    const hasAccess = r.can_request || r.can_view_all
    const roleNames = [
        r.admin && 'Admin',
        r.can_request && !r.admin && 'Requester',
        r.can_approve && 'Approver',
        r.can_process && !r.admin && 'Processor',
    ].filter(Boolean) as string[]

    return (
        <div className="page">
            <header className="header">
                <div>
                    <h1>NNAS Refunds</h1>
                    <div className="muted">
                        {me.you_are || 'Unknown user'}
                        {roleNames.map((n) => <span key={n} className="role">{n}</span>)}
                    </div>
                </div>
                {hasAccess && (
                    <nav className="tabs">
                        <button className={tab === 'requests' ? 'tab active' : 'tab'} onClick={() => setTab('requests')}>
                            Requests
                        </button>
                        {r.can_request && (
                            <button className={tab === 'new' ? 'tab active' : 'tab'} onClick={() => setTab('new')}>
                                New request
                            </button>
                        )}
                    </nav>
                )}
            </header>

            {me.test_mode && (
                <Notice
                    kind="warn"
                    text={`Test mode: only application ${me.test_application_ids.join(', ')} can be locked.${
                        me.test_solo ? ' Admin can run every step alone with a test-mode Stripe key.' : ''
                    }`}
                />
            )}

            {!hasAccess && <Notice kind="err" text="You do not have access to refunds. Contact IT." />}

            {hasAccess && tab === 'requests' && <RequestsView me={me} refreshKey={refreshKey} />}
            {hasAccess && tab === 'new' && r.can_request && (
                <NewRequest
                    onCreated={() => {
                        setRefreshKey((k) => k + 1)
                        setTab('requests')
                    }}
                />
            )}
        </div>
    )
}

// ============================================================ Requests list

function RequestsView({ me, refreshKey }: { me: Me; refreshKey: number }) {
    const [view, setView] = useState<'open' | 'closed' | 'all'>('open')
    const [rows, setRows] = useState<Row[]>([])
    const [loading, setLoading] = useState(false)
    const [err, setErr] = useState('')
    const [selectedId, setSelectedId] = useState<number | null>(null)
    const [find, setFind] = useState('')

    async function load() {
        setLoading(true)
        setErr('')
        try {
            setRows((await api.list_requests({ view })) ?? [])
        } catch (e) {
            setErr(errText(e))
        }
        setLoading(false)
    }

    useEffect(() => {
        load()
    }, [view, refreshKey])

    const needle = find.trim().toLowerCase()
    const shown = needle
        ? rows.filter((x) =>
              [x.reference_id, x.application_id, x.ref, x.requested_by]
                  .some((v) => String(v ?? '').toLowerCase().includes(needle))
          )
        : rows
    const selected = rows.find((x) => x.id === selectedId) ?? null

    return (
        <div className="layout">
            <section className="card list">
                <div className="toolbar">
                    <div className="seg">
                        {(['open', 'closed', 'all'] as const).map((v) => (
                            <button key={v} className={view === v ? 'seg-btn active' : 'seg-btn'} onClick={() => setView(v)}>
                                {v[0].toUpperCase() + v.slice(1)}
                            </button>
                        ))}
                    </div>
                    <input
                        className="find"
                        placeholder="Find by application number, RF number or person"
                        value={find}
                        onChange={(e) => setFind(e.target.value)}
                    />
                    <button onClick={load} disabled={loading}>{loading ? 'Loading...' : 'Refresh'}</button>
                </div>
                <Notice kind="err" text={err} />
                {shown.length === 0 && !loading && !err && (
                    <div className="muted pad">{needle ? 'Nothing matches that search.' : 'No requests.'}</div>
                )}
                {shown.length > 0 && (
                    <div className="table-wrap">
                        <table>
                            <thead>
                                <tr>
                                    <th>Ref</th>
                                    <th>Status</th>
                                    <th>Application</th>
                                    <th className="num">Amount</th>
                                    <th>Reason</th>
                                    <th>Requested by</th>
                                    <th>Requested</th>
                                </tr>
                            </thead>
                            <tbody>
                                {shown.map((x) => (
                                    <tr
                                        key={x.id}
                                        className={x.id === selectedId ? 'selected' : ''}
                                        onClick={() => setSelectedId(x.id)}
                                    >
                                        <td>
                                            {x.ref}
                                            {x.stripe_mode === 'test' && <span className="test-tag">TEST</span>}
                                        </td>
                                        <td><Badge status={x.status} /></td>
                                        <td>{x.reference_id || x.application_id}</td>
                                        <td className="num">{x.amount}</td>
                                        <td>{x.reason}</td>
                                        <td>{x.requested_by}</td>
                                        <td>{when(x.requested_at)}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </section>

            <section className="card detail">
                {selected ? (
                    <Detail key={selected.id + selected.status} row={selected} me={me} onChanged={load} />
                ) : (
                    <div className="muted pad">Select a request to see details and actions.</div>
                )}
            </section>
        </div>
    )
}

// ================================================================== Detail

function Detail({ row, me, onChanged }: { row: Row; me: Me; onChanged: () => void }) {
    const r = me.roles
    const [comment, setComment] = useState('')
    const [refundId, setRefundId] = useState('')
    const [armed, setArmed] = useState(false)
    const [busy, setBusy] = useState(false)
    const [ok, setOk] = useState('')
    const [err, setErr] = useState('')

    const canDecide = row.status === 'pending_approval' && (r.can_approve || (r.admin && me.test_solo))
    const canLock = (row.status === 'approved' || row.status === 'refunded') && r.can_process

    async function run(fn: () => Promise<any>) {
        setBusy(true)
        setOk('')
        setErr('')
        try {
            const res = await fn()
            const flagged = (v: any) => (typeof v === 'string' && /FAILED|SKIPPED/.test(v) ? v : '')
            const sent = typeof res?.alert === 'string' && res.alert.startsWith('alert sent') ? `(${res.alert})` : ''
            const extra = [res?.next, sent, flagged(res?.backup), flagged(res?.alert)]
                .filter(Boolean)
                .join(' ')
            setOk(`${res?.result ?? 'Done.'} ${extra}`.trim())
            onChanged()
        } catch (e) {
            setErr(errText(e))
        }
        setBusy(false)
        setArmed(false)
    }

    const decide = (decision: 'approve' | 'reject') =>
        run(() => api.decide_request({ request_id: row.id, decision, comment }))

    const lock = () =>
        run(() => api.complete_and_lock({ request_id: row.id, stripe_refund_id: refundId.trim() }))

    return (
        <div>
            <div className="detail-head">
                <h2>{row.ref}</h2>
                <Badge status={row.status} />
            </div>

            <dl className="facts">
                <dt>Application</dt><dd>{row.reference_id} ({row.application_id})</dd>
                <dt>Program</dt><dd>{row.program}</dd>
                <dt>Amount</dt><dd>{row.amount} ({row.refund_type})</dd>
                <dt>Stripe payment</dt><dd className="mono">{row.charge_id}</dd>
                <dt>Reason</dt><dd>{row.reason}</dd>
                <dt>Details</dt><dd className="pre">{row.reason_details}</dd>
                <dt>Requested</dt><dd>{row.requested_by}, {when(row.requested_at)}</dd>
                {row.decided_by && (
                    <>
                        <dt>Decision</dt>
                        <dd>{row.decided_by}, {when(row.decided_at)}{row.decision_comment ? `: ${row.decision_comment}` : ''}</dd>
                    </>
                )}
                {row.stripe_refund_id && (
                    <>
                        <dt>Stripe refund</dt>
                        <dd className="mono">{row.stripe_refund_id}, verified {when(row.refund_verified_at)}</dd>
                    </>
                )}
                {row.lock_completed_at && (<><dt>Locked</dt><dd>{when(row.lock_completed_at)}</dd></>)}
            </dl>

            {row.lock_error && <Notice kind="err" text={`Last lock attempt failed: ${row.lock_error}`} />}

            {canDecide && (
                <div className="action">
                    <h3>Decision</h3>
                    <textarea
                        rows={3}
                        placeholder="Comment (required to reject)"
                        value={comment}
                        onChange={(e) => setComment(e.target.value)}
                    />
                    <div className="row">
                        <button className="primary" disabled={busy} onClick={() => decide('approve')}>Approve</button>
                        <button className="danger" disabled={busy || !comment.trim()} onClick={() => decide('reject')}>Reject</button>
                    </div>
                </div>
            )}

            {canLock && (
                <div className="action">
                    <h3>Issue refund</h3>
                    {row.status === 'approved' ? (
                        <>
                            <p className="muted small">
                                1. Issue the refund of {row.amount} in Stripe on payment {row.charge_id}.
                                <br />2. Paste the Stripe refund id (re_...) and press Issue refund. This verifies
                                the refund, locks the application at Refunded - Locked and alerts the approver.
                            </p>
                            <input
                                className="mono"
                                placeholder="re_..."
                                value={refundId}
                                onChange={(e) => {
                                    setRefundId(e.target.value)
                                    setArmed(false)
                                }}
                            />
                        </>
                    ) : (
                        <p className="muted small">Refund is already verified. This retries the SMApply lock only.</p>
                    )}
                    <div className="row">
                        {!armed ? (
                            <button
                                className="primary"
                                disabled={busy || (row.status === 'approved' && !refundId.trim())}
                                onClick={() => setArmed(true)}
                            >
                                Issue refund
                            </button>
                        ) : (
                            <>
                                <button className="danger" disabled={busy} onClick={lock}>
                                    {busy ? 'Working...' : 'Confirm: verify refund and lock application'}
                                </button>
                                <button disabled={busy} onClick={() => setArmed(false)}>Cancel</button>
                            </>
                        )}
                    </div>
                </div>
            )}

            {busy && <div className="muted small">Working...</div>}
            <Notice kind="ok" text={ok} />
            <Notice kind="err" text={err} />
        </div>
    )
}

// ============================================================== New request
//
// Requesters do not have Stripe ids. They give an application number (and an
// email only if the payment was made under a different one), the server finds
// the Stripe payments for that email, and they pick one from the list.

function NewRequest({ onCreated }: { onCreated: () => void }) {
    const [application, setApplication] = useState('')
    const [email, setEmail] = useState('')
    const [found, setFound] = useState<any>(null)
    const [paymentId, setPaymentId] = useState('')
    const [refundType, setRefundType] = useState<'full' | 'partial'>('full')
    const [amount, setAmount] = useState('')
    const [reason, setReason] = useState(REASONS[0])
    const [details, setDetails] = useState('')
    const [preview, setPreview] = useState<any>(null)
    const [busy, setBusy] = useState('')
    const [ok, setOk] = useState('')
    const [err, setErr] = useState('')

    // Changing the application or email clears everything found for the old one.
    function editLookup(setter: (v: string) => void) {
        return (v: string) => {
            setter(v)
            setFound(null)
            setPaymentId('')
            setPreview(null)
            setOk('')
        }
    }

    // Any other edit invalidates the last check, so Submit always matches a fresh Check.
    function edit<T>(setter: (v: T) => void) {
        return (v: T) => {
            setter(v)
            setPreview(null)
            setOk('')
        }
    }

    async function findPayments() {
        setBusy('find')
        setErr('')
        setOk('')
        setFound(null)
        setPaymentId('')
        setPreview(null)
        try {
            const res = await api.find_payments({ application: application.trim(), email: email.trim() })
            setFound(res)
            const pickable = (res?.payments ?? []).filter((p: any) => p.selectable)
            if (pickable.length === 1) setPaymentId(pickable[0].payment_id)
        } catch (e) {
            setErr(errText(e))
        }
        setBusy('')
    }

    function args(dry_run: boolean) {
        return {
            application: application.trim(),
            payment_id: paymentId,
            email: email.trim(),
            refund_type: refundType,
            refund_amount: refundType === 'partial' ? amount.trim() : '',
            reason,
            reason_details: details,
            dry_run,
        }
    }

    async function check() {
        setBusy('check')
        setErr('')
        setOk('')
        setPreview(null)
        try {
            setPreview(await api.create_request(args(true)))
        } catch (e) {
            setErr(errText(e))
        }
        setBusy('')
    }

    async function submit() {
        setBusy('submit')
        setErr('')
        try {
            const res = await api.create_request(args(false))
            const note = typeof res?.alert === 'string' && /sent|FAILED|SKIPPED/.test(res.alert) ? ` (${res.alert})` : ''
            setOk(`${res?.result ?? 'Submitted.'}${note}`)
            setPreview(null)
            setFound(null)
            setPaymentId('')
            setApplication('')
            setEmail('')
            setAmount('')
            setDetails('')
            setTimeout(onCreated, 1500)
        } catch (e) {
            setErr(errText(e))
        }
        setBusy('')
    }

    const payments: any[] = found?.payments ?? []
    const ready = paymentId && details.trim().length >= 15 && (refundType === 'full' || amount.trim())

    return (
        <div className="layout">
            <section className="card">
                <h2>New refund request</h2>

                <h3>1. Find the payment</h3>
                <div className="form">
                    <label>Application number or SMApply ID
                        <input value={application} placeholder="e.g. NNAS-T-0000000032 or 35367543"
                            onChange={(e) => editLookup(setApplication)(e.target.value)} />
                    </label>
                    <label>Payment email (optional)
                        <input value={email} type="email" placeholder="Leave blank to use the applicant email on the application"
                            onChange={(e) => editLookup(setEmail)(e.target.value)} />
                        <span className="muted small">Only needed if the payment was made with a different email.</span>
                    </label>
                </div>
                <div className="row">
                    <button className="primary" disabled={!!busy || !application.trim()} onClick={findPayments}>
                        {busy === 'find' ? 'Searching Stripe...' : 'Find payments'}
                    </button>
                </div>

                {found && (
                    <>
                        <dl className="facts found">
                            <dt>Application</dt><dd>{found.application?.reference_id} ({found.application?.application_id})</dd>
                            <dt>Program</dt><dd>{found.application?.program}</dd>
                            <dt>Stage / status</dt><dd>{found.application?.stage} / {found.application?.status ?? 'none'}</dd>
                            <dt>Email used</dt>
                            <dd>
                                {found.email_used}, {found.email_source === 'entered' ? 'as entered' : 'from the application'}
                                {found.stripe_mode === 'test' && <span className="test-tag">STRIPE TEST MODE</span>}
                            </dd>
                        </dl>
                        {(found.notes ?? []).map((n: string) => <Notice key={n} kind="warn" text={n} />)}

                        {payments.length > 0 && (
                            <>
                                <h3>2. Choose the payment to refund</h3>
                                <div className="table-wrap">
                                    <table>
                                        <thead>
                                            <tr>
                                                <th></th>
                                                <th>Paid</th>
                                                <th className="num">Amount</th>
                                                <th className="num">Refunded</th>
                                                <th className="num">Refundable</th>
                                                <th>Description</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {payments.map((p) => (
                                                <tr
                                                    key={p.payment_id}
                                                    className={
                                                        (p.payment_id === paymentId ? 'selected ' : '') +
                                                        (p.selectable ? '' : 'disabled')
                                                    }
                                                    onClick={() => p.selectable && edit(setPaymentId)(p.payment_id)}
                                                >
                                                    <td>
                                                        <input
                                                            type="radio"
                                                            name="payment"
                                                            checked={p.payment_id === paymentId}
                                                            disabled={!p.selectable}
                                                            onChange={() => edit(setPaymentId)(p.payment_id)}
                                                        />
                                                    </td>
                                                    <td>{when(p.paid_on)}</td>
                                                    <td className="num">{p.amount}</td>
                                                    <td className="num">{p.refunded}</td>
                                                    <td className="num">{p.refundable}</td>
                                                    <td>
                                                        {p.open_request
                                                            ? `Already requested: ${p.open_request}`
                                                            : p.refundable_cents <= 0
                                                                ? 'Fully refunded'
                                                                : p.description}
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            </>
                        )}
                    </>
                )}

                {paymentId && (
                    <>
                        <h3>3. Refund details</h3>
                        <div className="form">
                            <label>Refund type
                                <select value={refundType} onChange={(e) => edit(setRefundType)(e.target.value as any)}>
                                    <option value="full">Full (everything still refundable)</option>
                                    <option value="partial">Partial</option>
                                </select>
                            </label>
                            {refundType === 'partial' && (
                                <label>Amount
                                    <input value={amount} inputMode="decimal" placeholder="e.g. 125.00"
                                        onChange={(e) => edit(setAmount)(e.target.value)} />
                                </label>
                            )}
                            <label>Reason
                                <select value={reason} onChange={(e) => edit(setReason)(e.target.value)}>
                                    {REASONS.map((x) => <option key={x}>{x}</option>)}
                                </select>
                            </label>
                            <label>Details
                                <textarea rows={4} value={details}
                                    placeholder="Why the refund is due. No card or banking details."
                                    onChange={(e) => edit(setDetails)(e.target.value)} />
                                <span className="muted small">{details.trim().length} / 15 characters minimum</span>
                            </label>
                        </div>
                        <div className="row">
                            <button disabled={!!busy || !ready} onClick={check}>
                                {busy === 'check' ? 'Checking...' : 'Check'}
                            </button>
                            <button className="primary" disabled={!!busy || !preview} onClick={submit}>
                                {busy === 'submit' ? 'Submitting...' : 'Submit request'}
                            </button>
                        </div>
                        <p className="muted small">Nothing is saved until Submit.</p>
                    </>
                )}

                <Notice kind="ok" text={ok} />
                <Notice kind="err" text={err} />
            </section>

            <section className="card detail">
                {preview ? (
                    <div>
                        <h2>Check passed</h2>
                        <dl className="facts">
                            <dt>Application</dt><dd>{preview.application?.reference_id} ({preview.application?.application_id})</dd>
                            <dt>Program</dt><dd>{preview.application?.program}</dd>
                            <dt>Stripe</dt><dd>{preview.payment?.stripe_mode === 'test' ? 'TEST MODE' : 'Live'}</dd>
                            <dt>Paid</dt><dd>{preview.payment?.original}, {when(preview.payment?.paid_on)}</dd>
                            <dt>Already refunded</dt><dd>{preview.payment?.already_refunded}</dd>
                            <dt>Refundable</dt><dd>{preview.payment?.refundable}</dd>
                            <dt>This request</dt><dd><strong>{preview.request?.amount}</strong> ({preview.request?.refund_type})</dd>
                        </dl>
                        {(preview.warnings ?? []).map((w: string) => <Notice key={w} kind="warn" text={w} />)}
                    </div>
                ) : (
                    <div className="muted pad">
                        Enter the application number and press Find payments. Pick the payment, fill in the
                        refund details and press Check. The verified summary appears here.
                    </div>
                )}
            </section>
        </div>
    )
}
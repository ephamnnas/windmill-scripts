import * as React from 'react'
import { useEffect, useMemo, useRef, useState } from 'react'
import type { FormEvent } from 'react'

import { backendAsync, waitJob } from './wmill'
import './index.css'

type Regulator = 'RN' | 'LPN'
type ApplicantLanguage = 'English' | 'French'
type RegulatoryBody =
  | 'College of Nursing of New Brunswick'
  | 'Association Of New Brunswick Licensed Practical Nurses'

type BatchState =
  | 'idle'
  | 'ready'
  | 'running'
  | 'stopping'
  | 'complete'
  | 'stopped'

type RowStatus =
  | 'ready'
  | 'invalid'
  | 'queued'
  | 'submitting'
  | 'running'
  | 'success'
  | 'warning'
  | 'error'
  | 'skipped'

type BatchSettings = {
  regulator: Regulator
  regulatory_body: RegulatoryBody
  send_invite: boolean
}

type ScriptResult = {
  ok?: boolean
  overall_status?: string
  onboarding_completed?: boolean
  user_id?: number
  application_id?: number
  email?: string
  full_name?: string
  regulator?: Regulator
  regulatory_body?: string
  language?: ApplicantLanguage
  province?: string
  pilot_program_value?: string
  pilot_program_on_user?: string | null
  province_on_user?: string | null
  eligibility_completed?: boolean | null
  task_status?: number | null
  invite_requested?: boolean
  invite_sent?: boolean
  invite_confirmed?: boolean
  invite_status?: string
  invite_message?: string
  invite_error?: string | null
  [key: string]: unknown
}

type ApplicantRow = {
  id: string
  sourceLine: number
  first_name: string
  last_name: string
  email: string
  language: ApplicantLanguage | null
  validationErrors: string[]
  status: RowStatus
  jobId?: string
  elapsedSeconds?: number
  result?: ScriptResult
  error?: string
}

const REGULATORY_BODIES: Array<{
  value: RegulatoryBody
  regulator: Regulator
  shortLabel: string
  fullLabel: string
}> = [
  {
    value: 'College of Nursing of New Brunswick',
    regulator: 'RN',
    shortLabel: 'Registered Nurse',
    fullLabel: 'College of Nursing of New Brunswick',
  },
  {
    value: 'Association Of New Brunswick Licensed Practical Nurses',
    regulator: 'LPN',
    shortLabel: 'Licensed Practical Nurse',
    fullLabel: 'Association of New Brunswick Licensed Practical Nurses',
  },
]

const INITIAL_SETTINGS: BatchSettings = {
  regulator: 'RN',
  regulatory_body: 'College of Nursing of New Brunswick',
  send_invite: true,
}

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

const STATUS_LABELS: Record<RowStatus, string> = {
  ready: 'Ready',
  invalid: 'Fix row',
  queued: 'Queued',
  submitting: 'Submitting',
  running: 'Processing',
  success: 'Completed',
  warning: 'Completed with warning',
  error: 'Failed',
  skipped: 'Not run',
}

function cleanCell(value: string): string {
  return value.replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ')
}

function looksLikeHeader(cells: string[]): boolean {
  const normalized = cells.slice(0, 3).map((cell) => cell.toLowerCase())

  return (
    normalized[0]?.includes('first') &&
    normalized[1]?.includes('last') &&
    normalized[2]?.includes('email')
  )
}

function parseExcelRows(value: string): ApplicantRow[] {
  const parsed: ApplicantRow[] = []
  const lines = value.split(/\r?\n/)
  const duplicateMap = new Map<string, number[]>()

  lines.forEach((rawLine, index) => {
    if (!rawLine.replace(/\u00a0/g, ' ').trim()) return

    const cells = rawLine.split('\t').map(cleanCell)
    while (cells.length > 0 && !cells[cells.length - 1]) cells.pop()

    if (parsed.length === 0 && looksLikeHeader(cells)) return

    const first_name = cells[0] ?? ''
    const last_name = cells[1] ?? ''
    const email = (cells[2] ?? '').toLowerCase()
    const validationErrors: string[] = []

    if (cells.length < 3) {
      validationErrors.push(
        'Expected three Excel columns: first name, last name and email.',
      )
    }

    if (!first_name) validationErrors.push('First name is missing.')
    if (!last_name) validationErrors.push('Last name is missing.')

    if (!email) {
      validationErrors.push('Email is missing.')
    } else if (!EMAIL_PATTERN.test(email)) {
      validationErrors.push('Email format is invalid.')
    }

    const row: ApplicantRow = {
      id: `line-${index + 1}-${email || parsed.length}`,
      sourceLine: index + 1,
      first_name,
      last_name,
      email,
      language: null,
      validationErrors,
      status: validationErrors.length > 0 ? 'invalid' : 'ready',
    }

    parsed.push(row)

    if (email) {
      const indexes = duplicateMap.get(email) ?? []
      indexes.push(parsed.length - 1)
      duplicateMap.set(email, indexes)
    }
  })

  duplicateMap.forEach((indexes, email) => {
    if (indexes.length < 2) return

    indexes.forEach((rowIndex) => {
      const row = parsed[rowIndex]
      row.validationErrors = [
        ...row.validationErrors,
        `Duplicate email in this batch: ${email}`,
      ]
      row.status = 'invalid'
    })
  })

  return parsed
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message
  if (typeof error === 'string') return error

  if (error && typeof error === 'object') {
    const record = error as Record<string, unknown>
    const candidates = [
      record.message,
      record.error,
      record.detail,
      record.reason,
    ]
    const readable = candidates.find((value) => typeof value === 'string')

    if (typeof readable === 'string') return readable

    try {
      return JSON.stringify(error, null, 2)
    } catch {
      return 'The Windmill job failed without a readable error message.'
    }
  }

  return 'The Windmill job failed without a readable error message.'
}

function formatElapsed(seconds: number): string {
  const minutes = Math.floor(seconds / 60)
  const remainingSeconds = seconds % 60

  return minutes > 0
    ? `${minutes}m ${remainingSeconds.toString().padStart(2, '0')}s`
    : `${remainingSeconds}s`
}

function pilotProgramFor(language: ApplicantLanguage | null): string {
  if (language === 'English') return 'NB-E'
  if (language === 'French') return 'NB-F'
  return 'Not assigned'
}

function StatusBadge({ status }: { status: RowStatus }) {
  return (
    <span className={`nnas-batch-status nnas-batch-status--${status}`}>
      {status === 'submitting' || status === 'running' ? (
        <span className="nnas-batch-mini-spinner" aria-hidden="true" />
      ) : null}
      {STATUS_LABELS[status]}
    </span>
  )
}

function LanguageSelector({
  value,
  disabled,
  onChange,
}: {
  value: ApplicantLanguage | null
  disabled: boolean
  onChange: (language: ApplicantLanguage) => void
}) {
  return (
    <div
      className="nnas-batch-row-language"
      role="group"
      aria-label="Applicant language"
    >
      <button
        type="button"
        className={value === 'English' ? 'is-selected' : ''}
        disabled={disabled}
        onClick={() => onChange('English')}
        aria-pressed={value === 'English'}
      >
        <strong>EN</strong>
        <small>NB-E</small>
      </button>

      <button
        type="button"
        className={value === 'French' ? 'is-selected' : ''}
        disabled={disabled}
        onClick={() => onChange('French')}
        aria-pressed={value === 'French'}
      >
        <strong>FR</strong>
        <small>NB-F</small>
      </button>
    </div>
  )
}

export default function App() {
  const [pasteValue, setPasteValue] = useState('')
  const [rows, setRows] = useState<ApplicantRow[]>([])
  const [settings, setSettings] = useState<BatchSettings>(INITIAL_SETTINGS)
  const [batchState, setBatchState] = useState<BatchState>('idle')
  const [batchStartedAt, setBatchStartedAt] = useState<number | null>(null)
  const [batchElapsed, setBatchElapsed] = useState(0)
  const [listMessage, setListMessage] = useState<string | null>(null)

  const stopRequested = useRef(false)
  const activeBatch = useRef(0)

  const isBusy = batchState === 'running' || batchState === 'stopping'

  const availableBodies = useMemo(
    () =>
      REGULATORY_BODIES.filter(
        (body) => body.regulator === settings.regulator,
      ),
    [settings.regulator],
  )

  const summary = useMemo(() => {
    const count = (status: RowStatus) =>
      rows.filter((row) => row.status === status).length

    return {
      total: rows.length,
      invalid: count('invalid'),
      unassigned: rows.filter(
        (row) => row.validationErrors.length === 0 && !row.language,
      ).length,
      english: rows.filter((row) => row.language === 'English').length,
      french: rows.filter((row) => row.language === 'French').length,
      queued: count('queued'),
      active: count('submitting') + count('running'),
      success: count('success'),
      warning: count('warning'),
      error: count('error'),
      skipped: count('skipped'),
      processed:
        count('success') +
        count('warning') +
        count('error') +
        count('skipped'),
    }
  }, [rows])

  const currentRow = rows.find(
    (row) => row.status === 'submitting' || row.status === 'running',
  )

  const progressPercent =
    summary.total > 0
      ? Math.round((summary.processed / summary.total) * 100)
      : 0

  const canRun =
    rows.length > 0 && summary.invalid === 0 && summary.unassigned === 0

  useEffect(() => {
    if (!isBusy || batchStartedAt === null) return

    const updateElapsed = () => {
      setBatchElapsed(
        Math.max(0, Math.floor((Date.now() - batchStartedAt) / 1000)),
      )
    }

    updateElapsed()
    const timer = window.setInterval(updateElapsed, 1000)

    return () => window.clearInterval(timer)
  }, [isBusy, batchStartedAt])

  function updateSetting<K extends keyof BatchSettings>(
    field: K,
    value: BatchSettings[K],
  ) {
    setSettings((current) => ({
      ...current,
      [field]: value,
    }))
  }

  function changeRegulator(regulator: Regulator) {
    const body = REGULATORY_BODIES.find(
      (item) => item.regulator === regulator,
    )

    if (!body) return

    setSettings((current) => ({
      ...current,
      regulator,
      regulatory_body: body.value,
    }))
  }

  function updateRow(rowId: string, patch: Partial<ApplicantRow>) {
    setRows((current) =>
      current.map((row) =>
        row.id === rowId
          ? {
              ...row,
              ...patch,
            }
          : row,
      ),
    )
  }

  function assignRowLanguage(
    rowId: string,
    language: ApplicantLanguage,
  ) {
    updateRow(rowId, { language })
  }

  function assignAllLanguages(language: ApplicantLanguage) {
    setRows((current) =>
      current.map((row) => {
        if (row.validationErrors.length > 0) return row

        return {
          ...row,
          language,
        }
      }),
    )

    setListMessage(
      `Assigned ${language} to all valid applicants. ${
        language === 'English' ? 'NB-E' : 'NB-F'
      } will be written to Pilot Program field 153504.`,
    )
  }

  function clearLanguageAssignments() {
    setRows((current) =>
      current.map((row) => ({
        ...row,
        language: null,
      })),
    )

    setListMessage('Language assignments cleared.')
  }

  function loadApplicants() {
    const parsed = parseExcelRows(pasteValue)

    setRows(parsed)
    setBatchState(parsed.length > 0 ? 'ready' : 'idle')
    setBatchStartedAt(null)
    setBatchElapsed(0)

    if (parsed.length === 0) {
      setListMessage('Paste at least one non-empty Excel row.')
      return
    }

    const invalid = parsed.filter(
      (row) => row.validationErrors.length > 0,
    ).length

    if (invalid > 0) {
      setListMessage(
        `${parsed.length} rows loaded. Fix or remove ${invalid} invalid row${
          invalid === 1 ? '' : 's'
        }, then assign English or French to each valid applicant.`,
      )
      return
    }

    setListMessage(
      `${parsed.length} applicant${
        parsed.length === 1 ? '' : 's'
      } loaded. Assign English or French to each applicant below.`,
    )
  }

  function removeRow(rowId: string) {
    setRows((current) => current.filter((row) => row.id !== rowId))
  }

  function clearBatch() {
    activeBatch.current += 1
    stopRequested.current = true
    setPasteValue('')
    setRows([])
    setBatchState('idle')
    setBatchStartedAt(null)
    setBatchElapsed(0)
    setListMessage(null)
  }

  function requestStop() {
    stopRequested.current = true
    setBatchState('stopping')
  }

  async function runBatch(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (rows.length === 0) {
      setListMessage('Load applicants from the pasted Excel rows first.')
      return
    }

    if (summary.invalid > 0) {
      setListMessage(
        'Fix or remove all invalid rows before starting the batch.',
      )
      return
    }

    if (summary.unassigned > 0) {
      setListMessage(
        `Assign English or French to all applicants. ${summary.unassigned} applicant${
          summary.unassigned === 1 ? ' is' : 's are'
        } still unassigned.`,
      )
      return
    }

    const batchToken = activeBatch.current + 1
    activeBatch.current = batchToken
    stopRequested.current = false

    const preparedRows = rows.map((row) => ({
      ...row,
      status: 'queued' as RowStatus,
      jobId: undefined,
      result: undefined,
      error: undefined,
      elapsedSeconds: undefined,
    }))

    setRows(preparedRows)
    setBatchState('running')
    setBatchStartedAt(Date.now())
    setBatchElapsed(0)
    setListMessage(null)

    for (let index = 0; index < preparedRows.length; index += 1) {
      if (activeBatch.current !== batchToken) return

      const row = preparedRows[index]

      if (stopRequested.current) {
        setRows((current) =>
          current.map((item) =>
            item.status === 'queued'
              ? {
                  ...item,
                  status: 'skipped',
                }
              : item,
          ),
        )

        setBatchState('stopped')
        return
      }

      if (!row.language) {
        updateRow(row.id, {
          status: 'error',
          error:
            'No language was assigned. Select English or French before running.',
        })
        continue
      }

      const applicantLanguage = row.language
      const rowStartedAt = Date.now()

      updateRow(row.id, {
        status: 'submitting',
      })

      try {
        const jobId = await backendAsync.onboard_applicant({
          email: row.email,
          first_name: row.first_name,
          last_name: row.last_name,
          regulator: settings.regulator,
          regulatory_body: settings.regulatory_body,
          language: applicantLanguage,
          send_invite: settings.send_invite,
        })

        if (activeBatch.current !== batchToken) return

        updateRow(row.id, {
          status: 'running',
          jobId: String(jobId),
        })

        const response = (await waitJob(jobId)) as ScriptResult

        if (activeBatch.current !== batchToken) return

        if (
          !response ||
          typeof response !== 'object' ||
          response.onboarding_completed !== true ||
          response.user_id == null ||
          response.application_id == null
        ) {
          throw new Error(
            `The script returned an incomplete result:\n${JSON.stringify(
              response,
              null,
              2,
            )}`,
          )
        }

        const hasWarning =
          response.overall_status === 'completed_invitation_failed' ||
          (settings.send_invite && response.invite_confirmed !== true)

        updateRow(row.id, {
          status: hasWarning ? 'warning' : 'success',
          result: response,
          elapsedSeconds: Math.max(
            0,
            Math.floor((Date.now() - rowStartedAt) / 1000),
          ),
        })
      } catch (caughtError) {
        if (activeBatch.current !== batchToken) return

        updateRow(row.id, {
          status: 'error',
          error: getErrorMessage(caughtError),
          elapsedSeconds: Math.max(
            0,
            Math.floor((Date.now() - rowStartedAt) / 1000),
          ),
        })
      }
    }

    if (activeBatch.current === batchToken) {
      setBatchState('complete')
    }
  }

  return (
    <main className="nnas-batch-app">
      <div className="nnas-batch-shell">
        <header className="nnas-batch-header">
          <div>
            <div className="nnas-batch-eyebrow">
              NNAS • SMApply administration
            </div>
            <h1>Batch expedited applicant onboarding</h1>
            <p>
              Paste applicants from Excel, load the list, then assign English
              or French to each applicant before submitting the batch.
            </p>
          </div>

          <div className="nnas-batch-language-summary">
            <div>
              <span>English</span>
              <strong>{summary.english}</strong>
              <small>NB-E</small>
            </div>
            <div>
              <span>French</span>
              <strong>{summary.french}</strong>
              <small>NB-F</small>
            </div>
            <div className={summary.unassigned > 0 ? 'has-unassigned' : ''}>
              <span>Unassigned</span>
              <strong>{summary.unassigned}</strong>
              <small>Required</small>
            </div>
          </div>
        </header>

        <form
          className="nnas-batch-layout"
          onSubmit={runBatch}
          noValidate
        >
          <div className="nnas-batch-main-column">
            <section className="nnas-batch-card">
              <div className="nnas-batch-card-heading">
                <span>1</span>
                <div>
                  <h2>Batch settings</h2>
                  <p>
                    Profession, regulatory body and invitation behaviour apply
                    to every applicant. Language is assigned per applicant
                    after loading.
                  </p>
                </div>
              </div>

              <div className="nnas-batch-settings-grid">
                <fieldset
                  className="nnas-batch-fieldset"
                  disabled={isBusy}
                >
                  <legend>Nursing profession</legend>

                  <div className="nnas-batch-segmented">
                    {REGULATORY_BODIES.map((body) => (
                      <label
                        key={body.regulator}
                        className={
                          settings.regulator === body.regulator
                            ? 'is-selected'
                            : ''
                        }
                      >
                        <input
                          type="radio"
                          name="regulator"
                          value={body.regulator}
                          checked={
                            settings.regulator === body.regulator
                          }
                          onChange={() =>
                            changeRegulator(body.regulator)
                          }
                        />
                        <strong>{body.regulator}</strong>
                        <small>{body.shortLabel}</small>
                      </label>
                    ))}
                  </div>
                </fieldset>

                <label className="nnas-batch-field">
                  <span>Regulatory body</span>
                  <select
                    value={settings.regulatory_body}
                    disabled={isBusy}
                    onChange={(event) =>
                      updateSetting(
                        'regulatory_body',
                        event.target.value as RegulatoryBody,
                      )
                    }
                  >
                    {availableBodies.map((body) => (
                      <option
                        key={body.value}
                        value={body.value}
                      >
                        {body.fullLabel}
                      </option>
                    ))}
                  </select>
                </label>
              </div>

              <label className="nnas-batch-switch-row nnas-batch-switch-row--full">
                <span>
                  <strong>Send invitation emails</strong>
                  <small>
                    Send each invitation after that applicant’s onboarding flow
                    completes.
                  </small>
                </span>
                <input
                  type="checkbox"
                  role="switch"
                  checked={settings.send_invite}
                  disabled={isBusy}
                  onChange={(event) =>
                    updateSetting(
                      'send_invite',
                      event.target.checked,
                    )
                  }
                />
              </label>
            </section>

            <section className="nnas-batch-card">
              <div className="nnas-batch-card-heading">
                <span>2</span>
                <div>
                  <h2>Paste from Excel</h2>
                  <p>
                    Copy three columns in this order: first name, last name and
                    email. A header row is optional.
                  </p>
                </div>
              </div>

              <label className="nnas-batch-paste-field">
                <span>Applicant rows</span>
                <textarea
                  value={pasteValue}
                  disabled={isBusy}
                  spellCheck={false}
                  onChange={(event) =>
                    setPasteValue(event.target.value)
                  }
                  placeholder={
                    'Malha\tBennedine\tlizabennedine@gmail.com\n' +
                    'Christine Zoe\tLaniog\tlaniog.christine@gmail.com\n' +
                    'Salvacion\tJavier\tshannyl28@yahoo.com\n' +
                    'Cecile\tBiclar\tnarsces35@gmail.com'
                  }
                />
              </label>

              <div className="nnas-batch-paste-actions">
                <button
                  className="nnas-batch-button nnas-batch-button--secondary"
                  type="button"
                  disabled={isBusy || !pasteValue.trim()}
                  onClick={loadApplicants}
                >
                  Load applicants
                </button>

                <button
                  className="nnas-batch-button nnas-batch-button--ghost"
                  type="button"
                  disabled={isBusy && rows.length === 0}
                  onClick={clearBatch}
                >
                  Clear
                </button>

                {listMessage ? (
                  <span
                    className={
                      summary.invalid > 0 ||
                      summary.unassigned > 0
                        ? 'is-warning'
                        : ''
                    }
                  >
                    {listMessage}
                  </span>
                ) : null}
              </div>
            </section>

            {rows.length > 0 ? (
              <section className="nnas-batch-card nnas-batch-table-card">
                <div className="nnas-batch-card-heading nnas-batch-card-heading--table">
                  <span>3</span>
                  <div>
                    <h2>Assign language and review queue</h2>
                    <p>
                      Select EN or FR for every applicant. Jobs then run
                      sequentially with a separate Windmill result for each row.
                    </p>
                  </div>

                  <strong className="nnas-batch-count">
                    {rows.length}
                  </strong>
                </div>

                <div className="nnas-batch-language-tools">
                  <span>Set valid applicants:</span>

                  <button
                    type="button"
                    disabled={isBusy}
                    onClick={() =>
                      assignAllLanguages('English')
                    }
                  >
                    All English
                    <small>NB-E</small>
                  </button>

                  <button
                    type="button"
                    disabled={isBusy}
                    onClick={() =>
                      assignAllLanguages('French')
                    }
                  >
                    All French
                    <small>NB-F</small>
                  </button>

                  <button
                    type="button"
                    className="is-clear"
                    disabled={isBusy}
                    onClick={clearLanguageAssignments}
                  >
                    Clear assignments
                  </button>

                  <div className="nnas-batch-language-counts">
                    <span>
                      EN <strong>{summary.english}</strong>
                    </span>
                    <span>
                      FR <strong>{summary.french}</strong>
                    </span>
                    <span
                      className={
                        summary.unassigned > 0
                          ? 'has-unassigned'
                          : ''
                      }
                    >
                      Unassigned{' '}
                      <strong>{summary.unassigned}</strong>
                    </span>
                  </div>
                </div>

                <div className="nnas-batch-table-wrap">
                  <table className="nnas-batch-table">
                    <thead>
                      <tr>
                        <th>#</th>
                        <th>Applicant</th>
                        <th>Email</th>
                        <th>Language</th>
                        <th>Status</th>
                        <th>Result</th>
                        <th aria-label="Actions" />
                      </tr>
                    </thead>

                    <tbody>
                      {rows.map((row, index) => {
                        const languageDisabled =
                          isBusy ||
                          row.status === 'success' ||
                          row.status === 'warning' ||
                          row.status === 'error'

                        return (
                          <tr
                            key={row.id}
                            className={`is-${row.status}`}
                          >
                            <td>{index + 1}</td>

                            <td>
                              <strong>
                                {row.first_name || 'Missing'}{' '}
                                {row.last_name}
                              </strong>
                              <small>
                                Source line {row.sourceLine}
                              </small>
                            </td>

                            <td className="nnas-batch-email-cell">
                              {row.email || '—'}
                            </td>

                            <td>
                              {row.validationErrors.length > 0 ? (
                                <span className="nnas-batch-muted">
                                  Fix row first
                                </span>
                              ) : (
                                <>
                                  <LanguageSelector
                                    value={row.language}
                                    disabled={languageDisabled}
                                    onChange={(language) =>
                                      assignRowLanguage(
                                        row.id,
                                        language,
                                      )
                                    }
                                  />
                                  <small className="nnas-batch-pilot-value">
                                    {pilotProgramFor(row.language)}
                                  </small>
                                </>
                              )}
                            </td>

                            <td>
                              <StatusBadge status={row.status} />
                              {row.elapsedSeconds != null ? (
                                <small>
                                  {formatElapsed(
                                    row.elapsedSeconds,
                                  )}
                                </small>
                              ) : null}
                            </td>

                            <td>
                              {row.validationErrors.length > 0 ? (
                                <ul className="nnas-batch-row-errors">
                                  {row.validationErrors.map((item) => (
                                    <li key={item}>{item}</li>
                                  ))}
                                </ul>
                              ) : row.error ? (
                                <details className="nnas-batch-row-details">
                                  <summary>View error</summary>
                                  <pre>{row.error}</pre>
                                </details>
                              ) : row.result ? (
                                <div className="nnas-batch-result-summary">
                                  <span>
                                    User{' '}
                                    <strong>
                                      {row.result.user_id}
                                    </strong>
                                  </span>
                                  <span>
                                    App{' '}
                                    <strong>
                                      {row.result.application_id}
                                    </strong>
                                  </span>
                                  <span>
                                    Pilot{' '}
                                    <strong>
                                      {row.result.pilot_program_value ??
                                        pilotProgramFor(
                                          row.language,
                                        )}
                                    </strong>
                                  </span>

                                  {row.result.invite_message ? (
                                    <small>
                                      {row.result.invite_message}
                                    </small>
                                  ) : null}

                                  <details className="nnas-batch-row-details">
                                    <summary>Full response</summary>
                                    <pre>
                                      {JSON.stringify(
                                        row.result,
                                        null,
                                        2,
                                      )}
                                    </pre>
                                  </details>
                                </div>
                              ) : row.jobId ? (
                                <span
                                  className="nnas-batch-job-id"
                                  title={row.jobId}
                                >
                                  Job {row.jobId}
                                </span>
                              ) : (
                                <span className="nnas-batch-muted">
                                  Waiting
                                </span>
                              )}
                            </td>

                            <td>
                              <button
                                type="button"
                                className="nnas-batch-remove-button"
                                disabled={isBusy}
                                onClick={() => removeRow(row.id)}
                                aria-label={`Remove ${row.first_name} ${row.last_name}`}
                              >
                                ×
                              </button>
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </section>
            ) : null}
          </div>

          <aside className="nnas-batch-side-column">
            <section className="nnas-batch-card nnas-batch-monitor">
              <div className="nnas-batch-card-heading nnas-batch-card-heading--compact">
                <span>4</span>
                <div>
                  <h2>Batch monitor</h2>
                  <p>
                    Live progress across the queued Windmill jobs.
                  </p>
                </div>
              </div>

              <div className="nnas-batch-progress-heading">
                <strong>
                  {summary.processed} of {summary.total || 0}
                </strong>
                <span>{progressPercent}%</span>
              </div>

              <div
                className="nnas-batch-progress-track"
                role="progressbar"
                aria-valuemin={0}
                aria-valuemax={summary.total || 1}
                aria-valuenow={summary.processed}
              >
                <span
                  style={{
                    width: `${progressPercent}%`,
                  }}
                />
              </div>

              <div className="nnas-batch-metrics">
                <div>
                  <span>Completed</span>
                  <strong>{summary.success}</strong>
                </div>
                <div>
                  <span>Warnings</span>
                  <strong>{summary.warning}</strong>
                </div>
                <div>
                  <span>Failed</span>
                  <strong>{summary.error}</strong>
                </div>
                <div>
                  <span>Elapsed</span>
                  <strong>{formatElapsed(batchElapsed)}</strong>
                </div>
              </div>

              <div className="nnas-batch-assignment-check">
                <div>
                  <span>English / NB-E</span>
                  <strong>{summary.english}</strong>
                </div>
                <div>
                  <span>French / NB-F</span>
                  <strong>{summary.french}</strong>
                </div>
                <div
                  className={
                    summary.unassigned > 0
                      ? 'has-unassigned'
                      : ''
                  }
                >
                  <span>Unassigned</span>
                  <strong>{summary.unassigned}</strong>
                </div>
              </div>

              {currentRow ? (
                <div className="nnas-batch-current">
                  <span>Currently processing</span>
                  <strong>
                    {currentRow.first_name}{' '}
                    {currentRow.last_name}
                  </strong>
                  <small>{currentRow.email}</small>
                  <small>
                    {currentRow.language} •{' '}
                    {pilotProgramFor(currentRow.language)}
                  </small>
                  {currentRow.jobId ? (
                    <code>{currentRow.jobId}</code>
                  ) : null}
                </div>
              ) : (
                <div className="nnas-batch-current nnas-batch-current--empty">
                  <span>Current status</span>
                  <strong>
                    {batchState === 'complete'
                      ? 'Batch complete'
                      : batchState === 'stopped'
                        ? 'Batch stopped'
                        : rows.length > 0
                          ? summary.unassigned > 0
                            ? 'Assign applicant languages'
                            : 'Ready to begin'
                          : 'No applicants loaded'}
                  </strong>
                </div>
              )}

              {!isBusy ? (
                <button
                  className="nnas-batch-button nnas-batch-button--primary"
                  type="submit"
                  disabled={!canRun}
                >
                  {summary.unassigned > 0
                    ? `Assign ${summary.unassigned} language${
                        summary.unassigned === 1 ? '' : 's'
                      }`
                    : `Run ${rows.length || 0} applicant${
                        rows.length === 1 ? '' : 's'
                      }`}
                </button>
              ) : (
                <button
                  className="nnas-batch-button nnas-batch-button--danger"
                  type="button"
                  disabled={batchState === 'stopping'}
                  onClick={requestStop}
                >
                  {batchState === 'stopping'
                    ? 'Stopping after current applicant…'
                    : 'Stop after current applicant'}
                </button>
              )}

              <div className="nnas-batch-monitor-note">
                Each row’s language is sent separately to the runnable:
                English writes <strong>NB-E</strong> and French writes{' '}
                <strong>NB-F</strong> to Pilot Program custom field{' '}
                <strong>153504</strong>.
              </div>
            </section>

            <section className="nnas-batch-card nnas-batch-help-card">
              <h3>Batch safeguards</h3>
              <ul>
                <li>
                  Every valid applicant must have an EN or FR assignment.
                </li>
                <li>
                  Duplicate emails in the pasted batch are blocked.
                </li>
                <li>
                  An exact email and user-ID match is required before inviting.
                </li>
                <li>
                  Invitation-only failures are displayed as warnings.
                </li>
                <li>
                  Stopped applicants remain marked as not run.
                </li>
              </ul>
            </section>
          </aside>
        </form>
      </div>
    </main>
  )
}

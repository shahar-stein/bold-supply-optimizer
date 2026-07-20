import {
  Activity,
  AlertCircle,
  ArrowRight,
  Download,
  DollarSign,
  Eye,
  FileSpreadsheet,
  Filter,
  Gauge,
  Globe,
  Layers,
  MousePointerClick,
  Sparkles,
  TrendingDown,
  TrendingUp,
  UploadCloud,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

import { Badge } from './components/ui/badge'
import { Button } from './components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from './components/ui/card'
import { Input } from './components/ui/input'

type ReportType = 'bundle' | 'country'
type ReasonType = 'Zero sRPM' | 'Low sRPM' | 'High requests / low cost share' | 'Very low fill rate'

type TrafficRow = {
  rowId: number
  supplyName: string
  dimension: string
  requests: number
  impressions: number
  cost: number
  srpm: number
}

type Thresholds = {
  lowSrpm: number
  highRequestShare: number
  lowCostShare: number
  lowFillRate: number
}

type Metrics = {
  requests: number
  impressions: number
  cost: number
}

type AnalyzedRow = TrafficRow & {
  groupKey: string
  fillRate: number
  requestShare: number
  costShare: number
  reasons: ReasonType[]
}

type GroupAnalysis = {
  groupKey: string
  before: Metrics
  after: Metrics
  badRows: AnalyzedRow[]
}

type SummaryMetrics = {
  before: Metrics
  after: Metrics
}

type OptimizationResult = {
  thresholds: Thresholds
  summary: SummaryMetrics
  score: number
  feasible: boolean
  badRows: number
  fillRateLift: number
  requestRetention: number
  impressionRetention: number
  costRetention: number
}

const DEFAULT_THRESHOLDS: Thresholds = {
  lowSrpm: 0.01,
  highRequestShare: 0.05,
  lowCostShare: 0.01,
  lowFillRate: 0.00005,
}

const REASON_LABELS: Record<ReasonType, string> = {
  'Zero sRPM': 'Zero',
  'Low sRPM': 'Low sRPM',
  'High requests / low cost share': 'Req > $',
  'Very low fill rate': 'Low fill',
}

const REPORT_CONFIG: Record<
  ReportType,
  { label: string; dimensionLabel: string; dimensionMatcher: RegExp; icon: LucideIcon }
> = {
  bundle: {
    label: 'Bundle',
    dimensionLabel: 'Platform ID',
    dimensionMatcher: /platformid|platform|bundle/,
    icon: Layers,
  },
  country: {
    label: 'Country',
    dimensionLabel: 'Country',
    dimensionMatcher: /country|geo/,
    icon: Globe,
  },
}

const formatNumber = (value: number) => new Intl.NumberFormat('en-US').format(Math.round(value))
const formatCost = (value: number) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 }).format(value)
const formatPercent = (value: number) => `${(value * 100).toFixed(2)}%`
const formatSignedPercent = (value: number) => `${value >= 0 ? '+' : ''}${(value * 100).toFixed(2)}%`

const metricFillRate = (m: Metrics) => (m.requests > 0 ? m.impressions / m.requests : 0)
const metricRetention = (after: number, before: number) => (before > 0 ? after / before : 0)

function parseNumber(value: unknown): number {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0
  }
  if (typeof value === 'string') {
    const normalized = value.replaceAll(',', '').trim()
    if (!normalized) {
      return 0
    }
    const parsed = Number(normalized)
    return Number.isFinite(parsed) ? parsed : 0
  }
  return 0
}

function normalizeHeader(value: unknown): string {
  return String(value ?? '').toLowerCase().replaceAll(/[^a-z0-9]/g, '')
}

function findHeaderIndex(headers: string[], matcher: RegExp): number {
  return headers.findIndex((value) => matcher.test(value))
}

function pickHeaderRow(rows: unknown[][], dimensionMatcher: RegExp): number {
  let winner = 0
  let bestScore = -1
  for (let i = 0; i < Math.min(10, rows.length); i += 1) {
    const header = rows[i].map(normalizeHeader)
    const score =
      Number(header.some((h) => /supplyname/.test(h))) +
      Number(header.some((h) => dimensionMatcher.test(h))) +
      Number(header.some((h) => /demandrequests|requests/.test(h))) +
      Number(header.some((h) => /supplyimpressions|impressions/.test(h))) +
      Number(header.some((h) => /cost/.test(h))) +
      Number(header.some((h) => /srpm/.test(h)))
    if (score > bestScore) {
      bestScore = score
      winner = i
    }
  }
  return winner
}

function extractRowsFromSheet(sheet: XLSX.WorkSheet, reportType: ReportType): { rows: TrafficRow[]; error?: string } {
  const config = REPORT_CONFIG[reportType]
  const data = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    blankrows: false,
    defval: null,
  }) as unknown[][]
  if (!data.length) {
    return { rows: [] }
  }

  const headerRowIndex = pickHeaderRow(data, config.dimensionMatcher)
  const headerRow = (data[headerRowIndex] ?? []).map(normalizeHeader)

  const supplyIndex = findHeaderIndex(headerRow, /supplyname|supply/)
  const dimensionIndex = findHeaderIndex(headerRow, config.dimensionMatcher)
  const requestsIndex = findHeaderIndex(headerRow, /demandrequests|requests/)
  const impressionsIndex = findHeaderIndex(headerRow, /supplyimpressions|impressions/)
  const costIndex = findHeaderIndex(headerRow, /^cost$|totalcost|spend/)
  const srpmIndex = findHeaderIndex(headerRow, /srpm/)

  if ([supplyIndex, dimensionIndex, requestsIndex, impressionsIndex, costIndex].some((idx) => idx < 0)) {
    return {
      rows: [],
      error: `Could not find required columns. Expected: Supply Name, ${config.dimensionLabel}, Supply Requests, Supply Impressions, Cost, sRPM $.`,
    }
  }

  const rows: TrafficRow[] = []
  let rowId = 1
  for (let i = headerRowIndex + 1; i < data.length; i += 1) {
    const row = data[i]
    if (!row) {
      continue
    }
    const supplyName = String(row[supplyIndex] ?? '').trim()
    const dimension = String(row[dimensionIndex] ?? '').trim()
    if (!supplyName || !dimension) {
      continue
    }

    const requests = parseNumber(row[requestsIndex])
    const impressions = parseNumber(row[impressionsIndex])
    const cost = parseNumber(row[costIndex])
    let srpm = srpmIndex >= 0 ? parseNumber(row[srpmIndex]) : 0
    if (srpm <= 0 && requests > 0 && cost > 0) {
      srpm = (cost / requests) * 1000
    }

    rows.push({
      rowId,
      supplyName,
      dimension,
      requests,
      impressions,
      cost,
      srpm,
    })
    rowId += 1
  }

  return { rows }
}

function buildAnalysis(rows: TrafficRow[], thresholds: Thresholds) {
  const totalsByGroup = new Map<string, Metrics>()
  for (const row of rows) {
    const groupKey = row.supplyName
    const existing = totalsByGroup.get(groupKey) ?? { requests: 0, impressions: 0, cost: 0 }
    existing.requests += row.requests
    existing.impressions += row.impressions
    existing.cost += row.cost
    totalsByGroup.set(groupKey, existing)
  }

  const groupMap = new Map<string, GroupAnalysis>()
  for (const row of rows) {
    const groupKey = row.supplyName
    const totals = totalsByGroup.get(groupKey)
    if (!totals) {
      continue
    }
    const requestShare = totals.requests > 0 ? row.requests / totals.requests : 0
    const costShare = totals.cost > 0 ? row.cost / totals.cost : 0
    const fillRate = row.requests > 0 ? row.impressions / row.requests : 0

    const reasons: ReasonType[] = []
    if (row.srpm === 0) {
      reasons.push('Zero sRPM')
    } else if (row.srpm > 0 && row.srpm < thresholds.lowSrpm) {
      reasons.push('Low sRPM')
    }
    if (requestShare >= thresholds.highRequestShare && costShare <= thresholds.lowCostShare) {
      reasons.push('High requests / low cost share')
    }
    if (fillRate <= thresholds.lowFillRate) {
      reasons.push('Very low fill rate')
    }

    const analysis = groupMap.get(groupKey) ?? {
      groupKey,
      before: { requests: 0, impressions: 0, cost: 0 },
      after: { requests: 0, impressions: 0, cost: 0 },
      badRows: [],
    }

    analysis.before.requests += row.requests
    analysis.before.impressions += row.impressions
    analysis.before.cost += row.cost

    const analyzedRow: AnalyzedRow = { ...row, groupKey, requestShare, costShare, fillRate, reasons }
    if (reasons.length > 0) {
      analysis.badRows.push(analyzedRow)
    } else {
      analysis.after.requests += row.requests
      analysis.after.impressions += row.impressions
      analysis.after.cost += row.cost
    }

    groupMap.set(groupKey, analysis)
  }

  const groups = Array.from(groupMap.values()).sort((a, b) => b.before.requests - a.before.requests)
  return groups
}

function summarizeGroups(groups: GroupAnalysis[]): SummaryMetrics {
  const before = groups.reduce(
    (acc, group) => ({
      requests: acc.requests + group.before.requests,
      impressions: acc.impressions + group.before.impressions,
      cost: acc.cost + group.before.cost,
    }),
    { requests: 0, impressions: 0, cost: 0 },
  )
  const after = groups.reduce(
    (acc, group) => ({
      requests: acc.requests + group.after.requests,
      impressions: acc.impressions + group.after.impressions,
      cost: acc.cost + group.after.cost,
    }),
    { requests: 0, impressions: 0, cost: 0 },
  )
  return { before, after }
}

function percentile(values: number[], point: number): number {
  if (!values.length) {
    return 0
  }
  const index = Math.min(values.length - 1, Math.max(0, Math.round((values.length - 1) * point)))
  return values[index]
}

function uniqueCandidates(values: number[], fallback: number[]): number[] {
  const normalized = [...values, ...fallback]
    .filter((value) => Number.isFinite(value) && value >= 0)
    .map((value) => Number(value.toPrecision(6)))
  return Array.from(new Set(normalized)).sort((a, b) => a - b)
}

function buildOptimizationCandidates(rows: TrafficRow[]): Thresholds[] {
  const totalsByGroup = new Map<string, Metrics>()
  for (const row of rows) {
    const groupKey = row.supplyName
    const existing = totalsByGroup.get(groupKey) ?? { requests: 0, impressions: 0, cost: 0 }
    existing.requests += row.requests
    existing.impressions += row.impressions
    existing.cost += row.cost
    totalsByGroup.set(groupKey, existing)
  }

  const positiveSrpm = rows.map((row) => row.srpm).filter((value) => value > 0).sort((a, b) => a - b)
  const fillRates = rows
    .map((row) => (row.requests > 0 ? row.impressions / row.requests : 0))
    .filter((value) => value > 0)
    .sort((a, b) => a - b)
  const requestShares: number[] = []
  const costShares: number[] = []

  for (const row of rows) {
    const groupKey = row.supplyName
    const totals = totalsByGroup.get(groupKey)
    if (!totals) {
      continue
    }
    if (totals.requests > 0) {
      requestShares.push(row.requests / totals.requests)
    }
    if (totals.cost > 0) {
      costShares.push(row.cost / totals.cost)
    }
  }

  requestShares.sort((a, b) => a - b)
  costShares.sort((a, b) => a - b)

  const lowSrpm = uniqueCandidates(
    [0, 0.001, DEFAULT_THRESHOLDS.lowSrpm, 0.025, ...[0.02, 0.05, 0.1, 0.15, 0.25].map((p) => percentile(positiveSrpm, p))],
    [DEFAULT_THRESHOLDS.lowSrpm],
  )
  const highRequestShare = uniqueCandidates(
    [0.01, 0.025, DEFAULT_THRESHOLDS.highRequestShare, 0.075, 0.1, ...[0.75, 0.85, 0.9, 0.95].map((p) => percentile(requestShares, p))],
    [DEFAULT_THRESHOLDS.highRequestShare],
  )
  const lowCostShare = uniqueCandidates(
    [0, 0.0025, 0.005, DEFAULT_THRESHOLDS.lowCostShare, 0.02, ...[0.05, 0.1, 0.2, 0.3].map((p) => percentile(costShares, p))],
    [DEFAULT_THRESHOLDS.lowCostShare],
  )
  const lowFillRate = uniqueCandidates(
    [0, DEFAULT_THRESHOLDS.lowFillRate, ...[0.01, 0.05, 0.1, 0.2, 0.3].map((p) => percentile(fillRates, p))],
    [DEFAULT_THRESHOLDS.lowFillRate],
  )

  const candidates: Thresholds[] = []
  for (const srpm of lowSrpm) {
    for (const requestShare of highRequestShare) {
      for (const costShare of lowCostShare) {
        for (const fillRate of lowFillRate) {
          candidates.push({
            lowSrpm: srpm,
            highRequestShare: requestShare,
            lowCostShare: costShare,
            lowFillRate: fillRate,
          })
        }
      }
    }
  }
  return candidates
}

function optimizeThresholds(rows: TrafficRow[]): OptimizationResult | null {
  if (!rows.length) {
    return null
  }

  let best: OptimizationResult | null = null
  let bestFeasible: OptimizationResult | null = null
  const candidates = buildOptimizationCandidates(rows)

  for (const thresholds of candidates) {
    const groups = buildAnalysis(rows, thresholds)
    const summary = summarizeGroups(groups)
    const beforeFillRate = metricFillRate(summary.before)
    const afterFillRate = metricFillRate(summary.after)
    const requestRetention = metricRetention(summary.after.requests, summary.before.requests)
    const impressionRetention = metricRetention(summary.after.impressions, summary.before.impressions)
    const costRetention = metricRetention(summary.after.cost, summary.before.cost)
    const fillRateLift = beforeFillRate > 0 ? (afterFillRate - beforeFillRate) / beforeFillRate : 0
    const badRows = groups.reduce((acc, group) => acc + group.badRows.length, 0)

    const feasible =
      badRows > 0 &&
      afterFillRate >= beforeFillRate &&
      requestRetention >= 0.35 &&
      impressionRetention >= 0.75 &&
      costRetention >= 0.7

    const retentionPenalty =
      Math.max(0, 0.35 - requestRetention) * 120 +
      Math.max(0, 0.75 - impressionRetention) * 100 +
      Math.max(0, 0.7 - costRetention) * 100

    const score =
      fillRateLift * 65 +
      impressionRetention * 18 +
      costRetention * 12 +
      requestRetention * 5 -
      retentionPenalty -
      (badRows === 0 ? 20 : 0)

    const result: OptimizationResult = {
      thresholds,
      summary,
      score,
      feasible,
      badRows,
      fillRateLift,
      requestRetention,
      impressionRetention,
      costRetention,
    }

    if (!best || result.score > best.score) {
      best = result
    }
    if (feasible && (!bestFeasible || result.score > bestFeasible.score)) {
      bestFeasible = result
    }
  }

  return bestFeasible ?? best
}

function App() {
  const [reportType, setReportType] = useState<ReportType>('bundle')
  const [workbooks, setWorkbooks] = useState<Record<ReportType, XLSX.WorkBook | null>>({ bundle: null, country: null })
  const [sheets, setSheets] = useState<Record<ReportType, string>>({ bundle: '', country: '' })
  const [uploadErrors, setUploadErrors] = useState<Record<ReportType, string>>({ bundle: '', country: '' })
  const [endpoints, setEndpoints] = useState<Record<ReportType, string>>({ bundle: 'ALL', country: 'ALL' })
  const [thresholds, setThresholds] = useState<Thresholds>(DEFAULT_THRESHOLDS)
  const [recommendedOptimization, setRecommendedOptimization] = useState<OptimizationResult | null>(null)

  const config = REPORT_CONFIG[reportType]
  const workbook = workbooks[reportType]
  const selectedSheet = sheets[reportType]
  const uploadError = uploadErrors[reportType]
  const selectedEndpoint = endpoints[reportType]

  const parsedSheet = useMemo(() => {
    if (!workbook || !selectedSheet) {
      return { rows: [] as TrafficRow[], error: '' }
    }
    const sheet = workbook.Sheets[selectedSheet]
    const { rows, error } = extractRowsFromSheet(sheet, reportType)
    return { rows, error: error ?? '' }
  }, [workbook, selectedSheet, reportType])

  const trafficRows = parsedSheet.rows
  const error = uploadError || parsedSheet.error

  const groups = useMemo(() => buildAnalysis(trafficRows, thresholds), [thresholds, trafficRows])

  const endpointOptions = useMemo(() => ['ALL', ...groups.map((group) => group.groupKey)], [groups])
  const filteredGroups = useMemo(
    () => (selectedEndpoint === 'ALL' ? groups : groups.filter((group) => group.groupKey === selectedEndpoint)),
    [groups, selectedEndpoint],
  )

  const summary = useMemo(() => summarizeGroups(filteredGroups), [filteredGroups])

  const allBadRows = useMemo(
    () =>
      filteredGroups
        .flatMap((group) => group.badRows)
        .sort((a, b) => b.requests - a.requests),
    [filteredGroups],
  )
  const badRows = useMemo(() => allBadRows.slice(0, 250), [allBadRows])

  useEffect(() => {
    if (!endpointOptions.includes(selectedEndpoint)) {
      setEndpoints((prev) => ({ ...prev, [reportType]: 'ALL' }))
    }
  }, [endpointOptions, selectedEndpoint, reportType])

  useEffect(() => {
    setRecommendedOptimization(null)
  }, [reportType, selectedEndpoint, trafficRows])

  function setEndpoint(value: string) {
    setEndpoints((prev) => ({ ...prev, [reportType]: value }))
  }

  function calculateRecommendedThresholds() {
    if (selectedEndpoint === 'ALL') {
      setRecommendedOptimization(null)
      return
    }

    const selectedRows = trafficRows.filter((row) => row.supplyName === selectedEndpoint)
    setRecommendedOptimization(optimizeThresholds(selectedRows))
  }

  function applyRecommendedThresholds() {
    if (!recommendedOptimization) {
      return
    }
    setThresholds(recommendedOptimization.thresholds)
  }

  function handleUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }
    const targetReport = reportType
    file
      .arrayBuffer()
      .then((buffer) => {
        const wb = XLSX.read(buffer, { type: 'array' })
        const firstDataSheet = wb.SheetNames.find((sheetName) => {
          const sheet = wb.Sheets[sheetName]
          const data = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false }) as unknown[][]
          return data.length > 1
        })
        setWorkbooks((prev) => ({ ...prev, [targetReport]: wb }))
        setSheets((prev) => ({ ...prev, [targetReport]: firstDataSheet ?? wb.SheetNames[0] ?? '' }))
        setEndpoints((prev) => ({ ...prev, [targetReport]: 'ALL' }))
        setUploadErrors((prev) => ({ ...prev, [targetReport]: '' }))
      })
      .catch(() => {
        setUploadErrors((prev) => ({ ...prev, [targetReport]: 'Could not parse file. Please upload a valid .xlsx report.' }))
      })
  }

  function exportFlaggedRows() {
    if (!allBadRows.length) {
      return
    }

    const exportedRows = allBadRows.map((row) => ({
      'Supply Name': row.supplyName,
      [config.dimensionLabel]: row.dimension,
      'Supply Requests': row.requests,
      'Supply Impressions': row.impressions,
      Cost: row.cost,
      'sRPM $': row.srpm,
      'Request Share': row.requestShare,
      'Cost Share': row.costShare,
      'Fill Rate': row.fillRate,
      Reasons: row.reasons.join(', '),
    }))

    const blockMap = new Map<
      string,
      {
        supplyName: string
        dimension: string
        requests: number
        impressions: number
        cost: number
      }
    >()

    for (const row of allBadRows) {
      const key = `${row.supplyName}:::${row.dimension}`
      const existing = blockMap.get(key) ?? {
        supplyName: row.supplyName,
        dimension: row.dimension,
        requests: 0,
        impressions: 0,
        cost: 0,
      }
      existing.requests += row.requests
      existing.impressions += row.impressions
      existing.cost += row.cost
      blockMap.set(key, existing)
    }

    const toBlock = Array.from(blockMap.values())
      .map((row) => {
        const fillRate = row.requests > 0 ? row.impressions / row.requests : 0
        const srpm = row.requests > 0 ? (row.cost / row.requests) * 1000 : 0
        return { supplyName: row.supplyName, dimension: row.dimension, requests: row.requests, fillRate, srpm }
      })
      .sort((a, b) => {
        const supplyDiff = a.supplyName.localeCompare(b.supplyName)
        if (supplyDiff !== 0) {
          return supplyDiff
        }
        const fillDiff = a.fillRate - b.fillRate
        if (fillDiff !== 0) {
          return fillDiff
        }
        const srpmDiff = a.srpm - b.srpm
        if (srpmDiff !== 0) {
          return srpmDiff
        }
        return b.requests - a.requests
      })
      .map((row) => {
        return {
          'Supply Name': row.supplyName,
          [config.dimensionLabel]: row.dimension,
        }
      })

    const worksheet = XLSX.utils.json_to_sheet(exportedRows)
    const blockWorksheet = XLSX.utils.json_to_sheet(toBlock)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Flagged Rows')
    XLSX.utils.book_append_sheet(workbook, blockWorksheet, `${config.dimensionLabel} to Block`.slice(0, 31))

    const selectedSegment = selectedEndpoint === 'ALL' ? 'all' : selectedEndpoint
    const safeSegment = selectedSegment.toLowerCase().replaceAll(/[^a-z0-9]+/g, '-').replaceAll(/^-|-$/g, '')
    XLSX.writeFile(workbook, `flagged-${reportType}-${safeSegment || 'rows'}.xlsx`)
  }

  return (
    <div className="mx-auto min-h-screen max-w-[1400px] px-4 py-6 md:px-8">
      <header className="mb-6 rounded-2xl border border-border bg-gradient-to-br from-primary/15 via-card to-card p-6 shadow-sm">
        <div className="flex flex-wrap items-center gap-3">
          <div className="rounded-xl bg-primary/15 p-2 text-primary">
            <FileSpreadsheet className="h-6 w-6" />
          </div>
          <div>
            <h1 className="text-2xl font-semibold tracking-tight md:text-3xl">Supply Traffic Optimizer</h1>
            <p className="text-sm text-muted-foreground md:text-base">
              Choose a report type, upload it, pick an endpoint, and spot bad traffic.
            </p>
          </div>
        </div>

        <div className="mt-5 grid grid-cols-2 gap-2 rounded-xl border border-border bg-card p-1 sm:max-w-md">
          {(Object.keys(REPORT_CONFIG) as ReportType[]).map((type) => {
            const TypeIcon = REPORT_CONFIG[type].icon
            const active = reportType === type
            return (
              <button
                key={type}
                type="button"
                onClick={() => setReportType(type)}
                className={`flex items-center justify-center gap-2 rounded-lg px-4 py-2.5 text-sm font-semibold transition-colors ${
                  active ? 'bg-primary text-primary-foreground shadow-sm' : 'text-muted-foreground hover:bg-secondary'
                }`}
              >
                <TypeIcon className="h-4 w-4" />
                {REPORT_CONFIG[type].label}
              </button>
            )
          })}
        </div>
      </header>

      <div className="grid min-w-0 gap-6 lg:grid-cols-[350px_minmax(0,1fr)]">
        <div className="min-w-0 space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <UploadCloud className="h-4 w-4 text-primary" />
                Upload {config.label} report
              </CardTitle>
              <CardDescription>
                Expected columns: Supply Name, {config.dimensionLabel}, Supply Requests, Supply Impressions, Cost, sRPM $.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <Input type="file" accept=".xlsx,.xls" onChange={handleUpload} />
              {workbook && (
                <div className="space-y-2">
                  <label className="text-sm font-medium" htmlFor="sheet">
                    Sheet
                  </label>
                  <select
                    id="sheet"
                    value={selectedSheet}
                    onChange={(event) => setSheets((prev) => ({ ...prev, [reportType]: event.target.value }))}
                    className="h-10 w-full rounded-md border border-input bg-background px-3 text-sm outline-none focus-visible:ring-2 focus-visible:ring-primary/30"
                  >
                    {workbook.SheetNames.map((sheetName) => (
                      <option key={sheetName} value={sheetName}>
                        {sheetName}
                      </option>
                    ))}
                  </select>
                </div>
              )}
              {workbook && endpointOptions.length > 1 && (
                <div className="space-y-2">
                  <label className="text-sm font-medium" htmlFor="endpoint">
                    Endpoint (Supply Name)
                  </label>
                  <select
                    id="endpoint"
                    value={selectedEndpoint}
                    onChange={(event) => setEndpoint(event.target.value)}
                    className="h-10 w-full rounded-md border border-input bg-background px-3 text-sm outline-none focus-visible:ring-2 focus-visible:ring-primary/30"
                  >
                    {endpointOptions.map((option) => (
                      <option key={option} value={option}>
                        {option === 'ALL' ? 'All endpoints' : option}
                      </option>
                    ))}
                  </select>
                </div>
              )}
              {error && (
                <div className="flex items-start gap-2 rounded-md border border-destructive/20 bg-destructive/10 p-3 text-sm text-destructive">
                  <AlertCircle className="mt-0.5 h-4 w-4" />
                  {error}
                </div>
              )}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Filter className="h-4 w-4 text-primary" />
                Optimization rules
              </CardTitle>
              <CardDescription>Adjust thresholds to control bad-traffic detection.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="rounded-lg border border-primary/25 bg-primary/5 p-3">
                <div className="flex items-start gap-2">
                  <div className="rounded-md bg-primary/15 p-1.5 text-primary">
                    <Sparkles className="h-4 w-4" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold">Auto optimize thresholds</p>
                    <p className="mt-1 text-xs text-muted-foreground">
                      Select one endpoint first, then calculate recommendations for that traffic only.
                    </p>
                  </div>
                </div>
                <Button
                  className="mt-3 w-full"
                  type="button"
                  onClick={calculateRecommendedThresholds}
                  disabled={selectedEndpoint === 'ALL' || trafficRows.length === 0}
                >
                  <Sparkles className="h-4 w-4" />
                  Calculate recommended params
                </Button>
                {selectedEndpoint === 'ALL' && trafficRows.length > 0 && (
                  <p className="mt-2 text-xs text-muted-foreground">
                    Choose a specific endpoint above to enable auto optimization.
                  </p>
                )}
                {recommendedOptimization && (
                  <div className="mt-3 grid gap-2 text-xs text-muted-foreground">
                    <div className="grid grid-cols-2 gap-2">
                      <span>Low sRPM: {recommendedOptimization.thresholds.lowSrpm.toFixed(6)}</span>
                      <span>Low fill: {formatPercent(recommendedOptimization.thresholds.lowFillRate)}</span>
                      <span>High req: {formatPercent(recommendedOptimization.thresholds.highRequestShare)}</span>
                      <span>Low cost: {formatPercent(recommendedOptimization.thresholds.lowCostShare)}</span>
                    </div>
                    <div className="flex flex-wrap gap-2 pt-1">
                      <Badge variant="secondary">Fill {formatSignedPercent(recommendedOptimization.fillRateLift)}</Badge>
                      <Badge variant="outline">Req kept {formatPercent(recommendedOptimization.requestRetention)}</Badge>
                      <Badge variant="outline">Imps kept {formatPercent(recommendedOptimization.impressionRetention)}</Badge>
                      <Badge variant="outline">Cost kept {formatPercent(recommendedOptimization.costRetention)}</Badge>
                    </div>
                    <Button className="w-full" type="button" variant="outline" onClick={applyRecommendedThresholds}>
                      Apply these params
                    </Button>
                    {!recommendedOptimization.feasible && (
                      <p className="text-destructive">No candidate met all guardrails; showing the best scored fallback.</p>
                    )}
                  </div>
                )}
              </div>
              <ThresholdInput
                label="Low sRPM threshold"
                value={thresholds.lowSrpm}
                onChange={(value) => setThresholds((prev) => ({ ...prev, lowSrpm: value }))}
              />
              <ThresholdInput
                label="High requests share (0-1)"
                value={thresholds.highRequestShare}
                onChange={(value) => setThresholds((prev) => ({ ...prev, highRequestShare: value }))}
              />
              <ThresholdInput
                label="Low cost share (0-1)"
                value={thresholds.lowCostShare}
                onChange={(value) => setThresholds((prev) => ({ ...prev, lowCostShare: value }))}
              />
              <ThresholdInput
                label="Low fill-rate threshold (0-1)"
                value={thresholds.lowFillRate}
                onChange={(value) => setThresholds((prev) => ({ ...prev, lowFillRate: value }))}
              />
            </CardContent>
          </Card>
        </div>

        <div className="min-w-0 space-y-6">
          <Card>
            <CardHeader>
              <CardTitle>Performance before vs after optimization</CardTitle>
              <CardDescription>After = same endpoint rows excluding rows flagged as bad traffic.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <PerformanceComparison summary={summary} />
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Per-endpoint optimization impact</CardTitle>
              <CardDescription>Sorted by supply requests.</CardDescription>
            </CardHeader>
            <CardContent className="overflow-x-auto">
              <table className="w-full min-w-[920px] text-left text-sm">
                <thead>
                  <tr className="border-b border-border text-muted-foreground">
                    <th className="py-2 pr-3 font-medium">Endpoint (Supply Name)</th>
                    <th className="py-2 pr-3 font-medium">Bad Rows</th>
                    <th className="py-2 pr-3 font-medium">Before Requests</th>
                    <th className="py-2 pr-3 font-medium">After Requests</th>
                    <th className="py-2 pr-3 font-medium">Before Cost</th>
                    <th className="py-2 pr-3 font-medium">After Cost</th>
                    <th className="py-2 pr-3 font-medium">Before FR</th>
                    <th className="py-2 pr-3 font-medium">After FR</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredGroups.map((group) => (
                    <tr key={group.groupKey} className="border-b border-border/70">
                      <td className="py-2 pr-3">{group.groupKey}</td>
                      <td className="py-2 pr-3">
                        <Badge variant={group.badRows.length > 0 ? 'destructive' : 'secondary'}>{group.badRows.length}</Badge>
                      </td>
                      <td className="py-2 pr-3">{formatNumber(group.before.requests)}</td>
                      <td className="py-2 pr-3">{formatNumber(group.after.requests)}</td>
                      <td className="py-2 pr-3">{formatCost(group.before.cost)}</td>
                      <td className="py-2 pr-3">{formatCost(group.after.cost)}</td>
                      <td className="py-2 pr-3">{formatPercent(metricFillRate(group.before))}</td>
                      <td className="py-2 pr-3">{formatPercent(metricFillRate(group.after))}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex-row items-start justify-between gap-4">
              <div>
                <CardTitle>Flagged bad traffic rows ({allBadRows.length})</CardTitle>
                <CardDescription>
                  Showing top {badRows.length} by supply requests. Export includes all flagged rows and a {config.dimensionLabel.toLowerCase()} block list.
                </CardDescription>
              </div>
              <Button type="button" variant="outline" size="sm" onClick={exportFlaggedRows} disabled={allBadRows.length === 0}>
                <Download className="h-4 w-4" />
                Export
              </Button>
            </CardHeader>
            <CardContent className="overflow-x-auto">
              <table className="w-full min-w-[1000px] text-left text-sm">
                <thead>
                  <tr className="border-b border-border text-muted-foreground">
                    <th className="py-2 pr-3 font-medium">Supply Name</th>
                    <th className="py-2 pr-3 font-medium">{config.dimensionLabel}</th>
                    <th className="py-2 pr-3 font-medium">Requests</th>
                    <th className="py-2 pr-3 font-medium">Impressions</th>
                    <th className="py-2 pr-3 font-medium">Cost</th>
                    <th className="py-2 pr-3 font-medium">sRPM</th>
                    <th className="py-2 pr-3 font-medium">Cost Share</th>
                    <th className="w-[180px] py-2 pr-3 font-medium">Reasons</th>
                  </tr>
                </thead>
                <tbody>
                  {badRows.map((row) => (
                    <tr key={`${row.groupKey}-${row.rowId}`} className="border-b border-border/70">
                      <td className="py-2 pr-3">{row.supplyName}</td>
                      <td className="py-2 pr-3">{row.dimension}</td>
                      <td className="py-2 pr-3">{formatNumber(row.requests)}</td>
                      <td className="py-2 pr-3">{formatNumber(row.impressions)}</td>
                      <td className="py-2 pr-3">{formatCost(row.cost)}</td>
                      <td className="py-2 pr-3">{row.srpm.toFixed(4)}</td>
                      <td className="py-2 pr-3">{formatPercent(row.costShare)}</td>
                      <td className="min-w-[180px] py-2 pr-3">
                        <div className="flex flex-wrap gap-1">
                          {row.reasons.map((reason) => (
                            <Badge
                              key={reason}
                              variant="secondary"
                              className="whitespace-nowrap rounded-md border border-border bg-secondary/70 px-2 py-0.5"
                              title={reason}
                            >
                              {REASON_LABELS[reason]}
                            </Badge>
                          ))}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  )
}

function ThresholdInput({
  label,
  value,
  onChange,
}: {
  label: string
  value: number
  onChange: (value: number) => void
}) {
  return (
    <div className="grid gap-2">
      <label className="text-sm font-medium">{label}</label>
      <Input
        type="number"
        step="0.0001"
        value={String(value)}
        onChange={(event) => onChange(parseNumber(event.target.value))}
      />
    </div>
  )
}

function metricChange(after: number, before: number) {
  if (before === 0) {
    return after > 0 ? 1 : 0
  }
  return (after - before) / before
}

function PerformanceComparison({ summary }: { summary: SummaryMetrics }) {
  return (
    <div className="grid gap-4 xl:grid-cols-2">
      <ComparisonMetric
        label="Supply Requests"
        icon={MousePointerClick}
        before={formatNumber(summary.before.requests)}
        after={formatNumber(summary.after.requests)}
        delta={metricChange(summary.after.requests, summary.before.requests)}
      />
      <ComparisonMetric
        label="Supply Impressions"
        icon={Eye}
        before={formatNumber(summary.before.impressions)}
        after={formatNumber(summary.after.impressions)}
        delta={metricChange(summary.after.impressions, summary.before.impressions)}
      />
      <ComparisonMetric
        label="Total Cost"
        icon={DollarSign}
        before={formatCost(summary.before.cost)}
        after={formatCost(summary.after.cost)}
        delta={metricChange(summary.after.cost, summary.before.cost)}
      />
      <ComparisonMetric
        label="Fill Rate"
        icon={Gauge}
        before={formatPercent(metricFillRate(summary.before))}
        after={formatPercent(metricFillRate(summary.after))}
        delta={metricChange(metricFillRate(summary.after), metricFillRate(summary.before))}
      />
    </div>
  )
}

function ComparisonMetric({
  label,
  icon: Icon,
  before,
  after,
  delta,
}: {
  label: string
  icon: LucideIcon
  before: string
  after: string
  delta: number
}) {
  const improved = delta > 0
  const declined = delta < 0
  const TrendIcon = improved ? TrendingUp : declined ? TrendingDown : Activity
  const trendClass = improved
    ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
    : declined
      ? 'border-rose-200 bg-rose-50 text-rose-700'
      : 'border-border bg-secondary text-muted-foreground'

  return (
    <div className="rounded-xl border border-border bg-secondary/25 p-4 shadow-sm">
      <div className="mb-4 flex items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <div className="rounded-lg bg-primary/15 p-2 text-primary">
            <Icon className="h-4 w-4" />
          </div>
          <p className="font-semibold">{label}</p>
        </div>
        <div className={`inline-flex items-center gap-1 rounded-full border px-2.5 py-1 text-xs font-semibold ${trendClass}`}>
          <TrendIcon className="h-3.5 w-3.5" />
          {formatSignedPercent(delta)}
        </div>
      </div>
      <div className="grid items-stretch gap-3 sm:grid-cols-[minmax(0,1fr)_auto_minmax(0,1fr)]">
        <div className="min-w-0 rounded-lg border border-border bg-card p-3">
          <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">Before</p>
          <p className="mt-2 break-words text-base font-bold tracking-tight">{before}</p>
        </div>
        <div className="flex items-center justify-center text-muted-foreground">
          <ArrowRight className="h-5 w-5" />
        </div>
        <div className="min-w-0 rounded-lg border border-primary/25 bg-primary/5 p-3">
          <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">After</p>
          <p className="mt-2 break-words text-base font-bold tracking-tight">{after}</p>
        </div>
      </div>
    </div>
  )
}

export default App

/**
 * Parse common date/time text into Excel 1900-system serials so number formats
 * (短日期 / 长日期 / Format Cells → Date) actually change the display.
 *
 * Univer's numfmt leaves string cells unchanged (`"2025-02-01"` stays that
 * text under every date pattern). AI-generated and imported sheets often store
 * dates as text; applying a date format must coerce them first.
 */

/** 'YYYY-MM-DD[ HH:mm[:ss]]' / slashes / Chinese 年月日 → Excel serial. */
export function textToExcelDateSerial(text: string): number | undefined {
  const trimmed = text.trim()
  if (!trimmed) return undefined

  const iso =
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(
      trimmed,
    )
  if (iso) {
    return partsToSerial(
      Number(iso[1]),
      Number(iso[2]),
      Number(iso[3]),
      Number(iso[4] ?? 0),
      Number(iso[5] ?? 0),
      Number(iso[6] ?? 0),
    )
  }

  const zh =
    /^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(
      trimmed,
    )
  if (zh) {
    return partsToSerial(
      Number(zh[1]),
      Number(zh[2]),
      Number(zh[3]),
      Number(zh[4] ?? 0),
      Number(zh[5] ?? 0),
      Number(zh[6] ?? 0),
    )
  }

  // m/d/yyyy or d/m/yyyy — only when the first number is > 12 (day-first) or
  // the second is > 12 (month-first). Ambiguous 1/2/2025 is left as text.
  const slash = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/.exec(trimmed)
  if (slash) {
    let month = Number(slash[1])
    let day = Number(slash[2])
    let year = Number(slash[3])
    if (year < 100) year += year >= 70 ? 1900 : 2000
    if (month > 12 && day <= 12) {
      const swap = month
      month = day
      day = swap
    } else if (month > 12 || day > 12) {
      if (month > 12) return undefined
    } else {
      return undefined
    }
    return partsToSerial(year, month, day, 0, 0, 0)
  }

  return undefined
}

/** 'H:MM[:SS]' → fraction of a day. */
export function textToExcelTimeFraction(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?(?:\s*(AM|PM))?$/i.exec(text.trim())
  if (!match) return undefined
  let hour = Number(match[1])
  const minute = Number(match[2])
  const second = Number(match[3] ?? 0)
  const ampm = match[4]?.toUpperCase()
  if (ampm === 'PM' && hour < 12) hour += 12
  if (ampm === 'AM' && hour === 12) hour = 0
  if (hour > 23 || minute > 59 || second > 59) return undefined
  return (hour * 3600 + minute * 60 + second) / 86_400
}

function partsToSerial(
  year: number,
  month: number,
  day: number,
  hour: number,
  minute: number,
  second: number,
): number | undefined {
  if (month < 1 || month > 12 || day < 1 || day > 31) return undefined
  if (hour > 23 || minute > 59 || second > 59) return undefined
  const days =
    (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86_400_000
  if (!Number.isFinite(days) || days < 0) return undefined
  const seconds = hour * 3600 + minute * 60 + second
  return seconds === 0 ? days : days + seconds / 86_400
}

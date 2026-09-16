import 'dotenv/config';

import { execFileSync } from 'node:child_process';

const checklist = process.env.CHECKLIST;
const previousWeeks = process.env.PREVIOUS_WEEKS;

if (!checklist || !previousWeeks) {
  throw new Error(
    'Missing CHECKLIST or PREVIOUS_WEEKS in .env',
  );
}

const toWindowsPath = (value: string) =>
  execFileSync(
    'wslpath',
    ['-w', value],
    { encoding: 'utf8' },
  ).trim();

const windowsChecklist = toWindowsPath(checklist);
const windowsPreviousWeeks = toWindowsPath(previousWeeks);

const powershellScript = `
$ErrorActionPreference = "Stop"

$checklist = '${windowsChecklist.replaceAll("'", "''")}'
$previousWeeks = '${windowsPreviousWeeks.replaceAll("'", "''")}'

$excel = $null
$workbook = $null

function Get-Monday([datetime]$date) {
    $offset = (([int]$date.DayOfWeek + 6) % 7)
    return $date.Date.AddDays(-$offset)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($checklist)

    $thisWeek = $workbook.Worksheets.Item("This Week")
    $history = $workbook.Worksheets.Item("Task History")
    $settings = $workbook.Worksheets.Item("Settings")

    $thisWeek.Columns.Item(5).NumberFormat = "mm/dd/yyyy"
    $thisWeek.Columns.Item(7).NumberFormat = "mm/dd/yyyy"
    $history.Columns.Item(4).NumberFormat = "mm/dd/yyyy"
    $history.Columns.Item(8).NumberFormat = "mm/dd/yyyy"

    $today = (Get-Date).Date
    $currentWeekStart = Get-Monday $today
    $currentWeekEnd = $currentWeekStart.AddDays(6)

    $title = [string]$thisWeek.Range("A1").Value2

    $match = [regex]::Match(
        $title,
        '([A-Za-z]{3} \\d{1,2}, \\d{4}) to ([A-Za-z]{3} \\d{1,2}, \\d{4})'
    )

    if (-not $match.Success) {
        throw "Could not read week dates from A1."
    }

    $sourceWeekStart =
        [datetime]::Parse($match.Groups[1].Value)

    $sourceWeekEnd =
        [datetime]::Parse($match.Groups[2].Value)

    $isNewWeek =
        $sourceWeekStart.Date -lt $currentWeekStart

    New-Item -ItemType Directory -Force -Path $previousWeeks |
        Out-Null

    if ($isNewWeek) {
        $archiveName =
            $sourceWeekStart.ToString("yyyy-MM-dd") +
            "_to_" +
            $sourceWeekEnd.ToString("yyyy-MM-dd") +
            ".xlsx"

        $archivePath =
            Join-Path $previousWeeks $archiveName

        if (Test-Path $archivePath) {
            Remove-Item $archivePath -Force
        }

        $workbook.SaveCopyAs($archivePath)
    }

    $archiveCompleted =
        ([string]$settings.Range("B6").Value2).Trim() -ieq "Yes"

    $carryOpen =
        ([string]$settings.Range("B7").Value2).Trim() -ieq "Yes"

    $preserveNotes =
        ([string]$settings.Range("B8").Value2).Trim() -ieq "Yes"

    $carriedRows = @()

    for ($row = 10; $row -le 209; $row++) {
        $task =
            ([string]$thisWeek.Cells.Item($row, 2).Value2).Trim()

        if (-not $task) {
            continue
        }

        $completed =
            ([string]$thisWeek.Cells.Item($row, 1).Value2).Trim() -eq "☑"

        if ($completed) {
            if ($archiveCompleted) {
                $historyRow = 4

                $history.Range(
                    $history.Cells.Item($historyRow, 1),
                    $history.Cells.Item($historyRow, 9)
                ).Insert(-4121) | Out-Null

                $source =
                    $thisWeek.Range(
                        $thisWeek.Cells.Item($row, 2),
                        $thisWeek.Cells.Item($row, 8)
                    )

                $destination =
                    $history.Range(
                        $history.Cells.Item($historyRow, 1),
                        $history.Cells.Item($historyRow, 7)
                    )

                $source.Copy() | Out-Null
                $destination.PasteSpecial(-4163) | Out-Null

                if (-not $preserveNotes) {
                    $history.Cells.Item($historyRow, 5).ClearContents()
                }

                $history.Cells.Item(
                    $historyRow,
                    8
                ).Value2 = [double]$today.ToOADate()

                $history.Cells.Item(
                    $historyRow,
                    9
                ).NumberFormat = "@"

                $history.Cells.Item(
                    $historyRow,
                    9
                ).Value2 =
                    $currentWeekStart.ToString("MM/dd/yyyy") +
                    " - " +
                    $currentWeekEnd.ToString("MM/dd/yyyy")
            }

            continue
        }

        if ($carryOpen) {
            $carriedRows += $row
        }
    }

    $targetRow = 10

    foreach ($sourceRow in $carriedRows) {
        if ($sourceRow -ne $targetRow) {
            $source =
                $thisWeek.Range(
                    $thisWeek.Cells.Item($sourceRow, 1),
                    $thisWeek.Cells.Item($sourceRow, 8)
                )

            $destination =
                $thisWeek.Range(
                    $thisWeek.Cells.Item($targetRow, 1),
                    $thisWeek.Cells.Item($targetRow, 8)
                )

            $source.Copy() | Out-Null
            $destination.PasteSpecial(-4163) | Out-Null
        }

        $thisWeek.Cells.Item($targetRow, 1).Value2 = "☐"

        if (-not $preserveNotes) {
            $thisWeek.Cells.Item($targetRow, 6).ClearContents()
        }

        $carriedValue =
            $thisWeek.Cells.Item($targetRow, 7).Value2

        $carriedFromDate = $null

        if (
            $null -ne $carriedValue -and
            ([string]$carriedValue).Trim()
        ) {
            if ($carriedValue -is [double]) {
                $carriedFromDate =
                    [datetime]::FromOADate(
                        [double]$carriedValue
                    )
            }
            else {
                $carriedFromDate =
                    [datetime]::Parse(
                        [string]$carriedValue
                    )
            }
        }

        if (
            -not $carriedFromDate -and
            $isNewWeek
        ) {
            $carriedFromDate =
                $sourceWeekStart.Date

            $thisWeek.Cells.Item(
                $targetRow,
                7
            ).Value2 =
                [double]$carriedFromDate.ToOADate()
        }

        $weeksCarried = 0

        if ($carriedFromDate) {
            $carriedWeekStart =
                Get-Monday $carriedFromDate

            $weeksCarried =
                [Math]::Max(
                    0,
                    [int][Math]::Floor(
                        (
                            $currentWeekStart -
                            $carriedWeekStart
                        ).TotalDays / 7
                    )
                )
        }

        $thisWeek.Cells.Item(
            $targetRow,
            8
        ).Value2 = [int]$weeksCarried

        $targetRow++
    }

    if ($targetRow -le 209) {
        $thisWeek.Range(
            $thisWeek.Cells.Item($targetRow, 1),
            $thisWeek.Cells.Item(209, 8)
        ).ClearContents()

        for ($row = $targetRow; $row -le 209; $row++) {
            $thisWeek.Cells.Item(
                $row,
                1
            ).Value2 = "☐"
        }
    }

    $thisWeek.Range("A1").Value2 =
        "Weekly Checklist — " +
        $currentWeekStart.ToString("MMM dd, yyyy") +
        " to " +
        $currentWeekEnd.ToString("MMM dd, yyyy")

    $workbook.Save()

    Write-Host (
        "Rollover complete. " +
        $carriedRows.Count +
        " task(s) carried forward."
    )
}
finally {
    if ($workbook) {
        $workbook.Close($false)
    }

    if ($excel) {
        $excel.Quit()
    }
}
`;

execFileSync(
  'powershell.exe',
  [
    '-NoProfile',
    '-Command',
    powershellScript,
  ],
  {
    stdio: 'inherit',
  },
);
import 'dotenv/config';

import { execFileSync } from 'node:child_process';

const checklist = process.env.CHECKLIST;

if (!checklist) {
  throw new Error(
    'CHECKLIST must be set in .env',
  );
}

const windowsChecklist = execFileSync(
  'wslpath',
  ['-w', checklist],
  { encoding: 'utf8' },
).trim();

const projectPath = process.cwd();

const taskName =
  'AutoTask Weekly Rollover';

const powershellScript = `
$ErrorActionPreference = "Stop"

$checklist = '${windowsChecklist.replaceAll("'", "''")}'

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open(
        $checklist,
        0,
        $true
    )

    $settings = $workbook.Worksheets.Item("Settings")

    $rolloverDay =
        ([string]$settings.Range("B5").Value2).Trim()

    $validDays = @(
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday"
    )

    $scheduledDay =
        $validDays |
        Where-Object {
            $_ -ieq $rolloverDay
        } |
        Select-Object -First 1

    if (-not $scheduledDay) {
        throw "Invalid rollover day: $rolloverDay"
    }

    $action = New-ScheduledTaskAction -Execute "wsl.exe" -Argument 'bash -lc "cd ${projectPath} && npm run rollover"'

    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $scheduledDay -At 8:00AM

    $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable

    Register-ScheduledTask -TaskName "${taskName}" -Action $action -Trigger $trigger -Settings $taskSettings -Force

    Write-Host ""
    Write-Host "AutoTask scheduled for every $scheduledDay at 8:00 AM."
    Write-Host "Missed runs will start when the PC becomes available again."
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
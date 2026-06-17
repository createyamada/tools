import json
import os
import subprocess
import tempfile

from openpyxl import Workbook
from openpyxl import load_workbook

RULE_FILE = "rule.json"
MASTER_FILE = "sender_master.xlsx"
FIRST_RUN_FILE = "first_run.flag"


def load_rules():

    if not os.path.exists(RULE_FILE):

        rules = {
            "[案件]": "案件管理",
            "[見積]": "見積管理",
            "[請求]": "請求管理",
            "[障害]": "障害対応"
        }

        with open(
            RULE_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                rules,
                f,
                ensure_ascii=False,
                indent=4
            )

    with open(
        RULE_FILE,
        "r",
        encoding="utf-8"
    ) as f:

        return json.load(f)


def load_sender_master():

    if not os.path.exists(MASTER_FILE):

        wb = Workbook()
        ws = wb.active

        ws["A1"] = "MailAddress"
        ws["B1"] = "FolderName"

        wb.save(MASTER_FILE)

    wb = load_workbook(MASTER_FILE)
    ws = wb.active

    sender_map = {}

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        mail = str(row[0]).strip().lower()

        if row[1]:
            folder = str(row[1]).strip()
        else:
            folder = mail

        sender_map[mail] = folder

    return wb, ws, sender_map


def save_new_sender(
    ws,
    sender_map,
    sender
):

    sender = sender.lower()

    if sender in sender_map:
        return

    ws.append([
        sender,
        sender
    ])

    sender_map[sender] = sender


rules = load_rules()

wb, ws, sender_map = load_sender_master()

is_first_run = not os.path.exists(
    FIRST_RUN_FILE
)

rules_json = json.dumps(
    rules,
    ensure_ascii=False
)

sender_json = json.dumps(
    sender_map,
    ensure_ascii=False
)

powershell_script = f"""
$isFirstRun = {'$true' if is_first_run else '$false'}

$rules = ConvertFrom-Json @'
{rules_json}
'@

$senderMap = ConvertFrom-Json @'
{sender_json}
'@

function Get-SenderAddress($mail)
{{
    try
    {{
        if($mail.SenderEmailType -eq "EX")
        {{
            $exchangeUser =
                $mail.Sender.GetExchangeUser()

            if($exchangeUser)
            {{
                return $exchangeUser.PrimarySmtpAddress
            }}
        }}

        return $mail.SenderEmailAddress
    }}
    catch
    {{
        return $mail.SenderEmailAddress
    }}
}}

$outlook =
    New-Object -ComObject Outlook.Application

$namespace =
    $outlook.GetNamespace("MAPI")

$inbox =
    $namespace.GetDefaultFolder(6)

$rootFolder =
    $inbox.Parent

$items = @()

if($isFirstRun)
{{
    foreach($item in $inbox.Items)
    {{
        $items += $item
    }}
}}
else
{{
    $unreadItems =
        $inbox.Items.Restrict(
            "[UnRead] = True"
        )

    foreach($item in $unreadItems)
    {{
        $items += $item
    }}
}}

foreach($mail in $items)
{{
    try
    {{
        if($mail.Class -ne 43)
        {{
            continue
        }}

        $subject = $mail.Subject

        foreach(
            $property
            in $rules.PSObject.Properties
        )
        {{
            $keyword =
                $property.Name

            $parentFolderName =
                $property.Value

            if(
                $subject -like
                "*$keyword*"
            )
            {{

                try
                {{
                    $parentFolder =
                        $rootFolder.Folders.Item(
                            $parentFolderName
                        )
                }}
                catch
                {{
                    $parentFolder =
                        $rootFolder.Folders.Add(
                            $parentFolderName
                        )
                }}

                $sender =
                    Get-SenderAddress $mail

                if(
                    [string]::IsNullOrWhiteSpace(
                        $sender
                    )
                )
                {{
                    $sender = "Unknown"
                }}

                $senderLower =
                    $sender.ToLower()

                $folderName = $sender

                if(
                    $senderMap.PSObject.Properties.Name `
                    -contains `
                    $senderLower
                )
                {{
                    $folderName =
                        $senderMap.$senderLower
                }}

                $folderName =
                    $folderName `
                    -replace
                    '[\\\\/:*?"<>|]',
                    '_'

                try
                {{
                    $targetFolder =
                        $parentFolder.Folders.Item(
                            $folderName
                        )
                }}
                catch
                {{
                    $targetFolder =
                        $parentFolder.Folders.Add(
                            $folderName
                        )
                }}

                $mail.Move(
                    $targetFolder
                )

                Write-Output(
                    "NEWMAIL|" +
                    $senderLower
                )

                break
            }}
        }}
    }}
    catch
    {{
        Write-Output(
            "ERROR|" +
            $_.Exception.Message
        )
    }}
}}
"""

with tempfile.NamedTemporaryFile(
    mode="w",
    suffix=".ps1",
    delete=False,
    encoding="utf-8"
) as f:

    f.write(
        powershell_script
    )

    ps_file = f.name

try:

    result = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            ps_file
        ],
        capture_output=True,
        text=True,
        encoding="utf-8"
    )

    for line in result.stdout.splitlines():

        if line.startswith(
            "NEWMAIL|"
        ):

            sender = (
                line.split(
                    "|",
                    1
                )[1]
                .strip()
                .lower()
            )

            save_new_sender(
                ws,
                sender_map,
                sender
            )

    wb.save(MASTER_FILE)

    if is_first_run:

        with open(
            FIRST_RUN_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            f.write("done")

finally:

    if os.path.exists(ps_file):

        os.remove(ps_file)

print("完了")import json
import os
import subprocess
import tempfile

from openpyxl import Workbook
from openpyxl import load_workbook

RULE_FILE = "rule.json"
MASTER_FILE = "sender_master.xlsx"
FIRST_RUN_FILE = "first_run.flag"


def load_rules():

    if not os.path.exists(RULE_FILE):

        rules = {
            "[案件]": "案件管理",
            "[見積]": "見積管理",
            "[請求]": "請求管理",
            "[障害]": "障害対応"
        }

        with open(
            RULE_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                rules,
                f,
                ensure_ascii=False,
                indent=4
            )

    with open(
        RULE_FILE,
        "r",
        encoding="utf-8"
    ) as f:

        return json.load(f)


def load_sender_master():

    if not os.path.exists(MASTER_FILE):

        wb = Workbook()
        ws = wb.active

        ws["A1"] = "MailAddress"
        ws["B1"] = "FolderName"

        wb.save(MASTER_FILE)

    wb = load_workbook(MASTER_FILE)
    ws = wb.active

    sender_map = {}

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        mail = str(row[0]).strip().lower()

        if row[1]:
            folder = str(row[1]).strip()
        else:
            folder = mail

        sender_map[mail] = folder

    return wb, ws, sender_map


def save_new_sender(
    ws,
    sender_map,
    sender
):

    sender = sender.lower()

    if sender in sender_map:
        return

    ws.append([
        sender,
        sender
    ])

    sender_map[sender] = sender


rules = load_rules()

wb, ws, sender_map = load_sender_master()

is_first_run = not os.path.exists(
    FIRST_RUN_FILE
)

rules_json = json.dumps(
    rules,
    ensure_ascii=False
)

sender_json = json.dumps(
    sender_map,
    ensure_ascii=False
)

powershell_script = f"""
$isFirstRun = {'$true' if is_first_run else '$false'}

$rules = ConvertFrom-Json @'
{rules_json}
'@

$senderMap = ConvertFrom-Json @'
{sender_json}
'@

function Get-SenderAddress($mail)
{{
    try
    {{
        if($mail.SenderEmailType -eq "EX")
        {{
            $exchangeUser =
                $mail.Sender.GetExchangeUser()

            if($exchangeUser)
            {{
                return $exchangeUser.PrimarySmtpAddress
            }}
        }}

        return $mail.SenderEmailAddress
    }}
    catch
    {{
        return $mail.SenderEmailAddress
    }}
}}

$outlook =
    New-Object -ComObject Outlook.Application

$namespace =
    $outlook.GetNamespace("MAPI")

$inbox =
    $namespace.GetDefaultFolder(6)

$rootFolder =
    $inbox.Parent

$items = @()

if($isFirstRun)
{{
    foreach($item in $inbox.Items)
    {{
        $items += $item
    }}
}}
else
{{
    $unreadItems =
        $inbox.Items.Restrict(
            "[UnRead] = True"
        )

    foreach($item in $unreadItems)
    {{
        $items += $item
    }}
}}

foreach($mail in $items)
{{
    try
    {{
        if($mail.Class -ne 43)
        {{
            continue
        }}

        $subject = $mail.Subject

        foreach(
            $property
            in $rules.PSObject.Properties
        )
        {{
            $keyword =
                $property.Name

            $parentFolderName =
                $property.Value

            if(
                $subject -like
                "*$keyword*"
            )
            {{

                try
                {{
                    $parentFolder =
                        $rootFolder.Folders.Item(
                            $parentFolderName
                        )
                }}
                catch
                {{
                    $parentFolder =
                        $rootFolder.Folders.Add(
                            $parentFolderName
                        )
                }}

                $sender =
                    Get-SenderAddress $mail

                if(
                    [string]::IsNullOrWhiteSpace(
                        $sender
                    )
                )
                {{
                    $sender = "Unknown"
                }}

                $senderLower =
                    $sender.ToLower()

                $folderName = $sender

                if(
                    $senderMap.PSObject.Properties.Name `
                    -contains `
                    $senderLower
                )
                {{
                    $folderName =
                        $senderMap.$senderLower
                }}

                $folderName =
                    $folderName `
                    -replace
                    '[\\\\/:*?"<>|]',
                    '_'

                try
                {{
                    $targetFolder =
                        $parentFolder.Folders.Item(
                            $folderName
                        )
                }}
                catch
                {{
                    $targetFolder =
                        $parentFolder.Folders.Add(
                            $folderName
                        )
                }}

                $mail.Move(
                    $targetFolder
                )

                Write-Output(
                    "NEWMAIL|" +
                    $senderLower
                )

                break
            }}
        }}
    }}
    catch
    {{
        Write-Output(
            "ERROR|" +
            $_.Exception.Message
        )
    }}
}}
"""

with tempfile.NamedTemporaryFile(
    mode="w",
    suffix=".ps1",
    delete=False,
    encoding="utf-8"
) as f:

    f.write(
        powershell_script
    )

    ps_file = f.name

try:

    result = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            ps_file
        ],
        capture_output=True,
        text=True,
        encoding="utf-8"
    )

    for line in result.stdout.splitlines():

        if line.startswith(
            "NEWMAIL|"
        ):

            sender = (
                line.split(
                    "|",
                    1
                )[1]
                .strip()
                .lower()
            )

            save_new_sender(
                ws,
                sender_map,
                sender
            )

    wb.save(MASTER_FILE)

    if is_first_run:

        with open(
            FIRST_RUN_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            f.write("done")

finally:

    if os.path.exists(ps_file):

        os.remove(ps_file)

print("完了")

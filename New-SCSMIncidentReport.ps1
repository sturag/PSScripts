function New-SCSMIncidentReport {
    <#
    .SYNOPSIS
    Generates a localized (sv|en) HTML report of active SCSM incidents and writes it to OutputPath.

    .DESCRIPTION
    Uses SCSM cmdlets to retrieve incidents with Status = Active, optionally filters by classification or tier/queue,
    sorts results, and builds a single-file UTF8 HTML report with a master table and expandable detail rows.
    UI strings and timestamp are localized based on the Language parameter (maps to SVE/ENU for SCSM lookups).

    .PARAMETER OutputPath
    Mandatory. File path where the generated HTML report will be saved.

    .PARAMETER Title
    Optional. Report title; defaults to a language-specific value when omitted.

    .PARAMETER OrderBy
    Sort column. Valid values: Id, CreatedDate, Title. Default: CreatedDate.

    .PARAMETER Descending
    Switch. If present, sort order is descending.

    .PARAMETER ClassificationLike
    Optional. Wildcard filter applied to Classification.DisplayName on the server side.

    .PARAMETER TierQueueLike
    Optional. Wildcard filter applied to TierQueue.DisplayName on the server side.

    .PARAMETER Language
    UI language and SCSM three-letter language mapping. Valid values: 'sv' (Swedish/SVE), 'en' (English/ENU). Default: 'sv'.

    .EXAMPLE
    New-SCSMIncidentReport -OutputPath C:\temp\incidents.html -Language en

    .REMARKS
    Requires SMLets: Get-SCSMClass, Get-SCSMObject, Get-SCSMRelationshipClass,
    Get-SCSMRelationshipObject, Get-SCSMEnumeration. Loads System.Web for HTML encoding.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$Title,
        [ValidateSet('Id', 'CreatedDate', 'Title')][string]$OrderBy = 'CreatedDate',
        [switch]$Descending,
        [string]$ClassificationLike,   # valfri server-side klassificeringsfilter
        [string]$TierQueueLike,
        [ValidateSet('sv', 'en')][string]$Language = 'sv'
    )

    try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue | Out-Null } catch {}
    $htmlEncode = { param($s) [System.Web.HttpUtility]::HtmlEncode([string]$s) }

    # map to SCSM three-letter language tag
    $threeLetter = @{ 'sv' = 'SVE'; 'en' = 'ENU' }
    $tl = $threeLetter[$Language]

    if (-not $Title) {
        $Title = if ($Language -eq 'sv') { 'Aktiva Incidenter' } else { 'Active Incidents' }
    }

    # localized UI strings (sv/en)
    $t = @{
        ID                 = @{ sv = 'ID'; en = 'ID' }
        Title              = @{ sv = 'Titel'; en = 'Title' }
        AffectedUser       = @{ sv = 'Berörd anv.'; en = 'Affected user' }
        AssignedTo         = @{ sv = 'Utförare'; en = 'Assigned to' }
        Created            = @{ sv = 'Skapad'; en = 'Created' }
        Status             = @{ sv = 'Status'; en = 'Status' }
        Classification     = @{ sv = 'Klassificering'; en = 'Classification' }
        Queue              = @{ sv = 'Supportgrupp'; en = 'TierQueue' }
        Related            = @{ sv = 'Relaterade'; en = 'Related' }
        GeneratedLabel     = @{ sv = 'Genererad'; en = 'Generated' }
        IncidentsCount     = @{ sv = '{0} ärenden • Genererad: {1}'; en = '{0} incidents • Generated: {1}' }
        ExpandAll          = @{ sv = 'Expandera alla'; en = 'Expand all' }
        CollapseAll        = @{ sv = 'Fäll ihop alla'; en = 'Collapse all' }
        AllClassifications = @{ sv = 'Alla klassificeringar'; en = 'All classifications' }
        AllCreators        = @{ sv = 'Alla skapare'; en = 'All creators' }
        AllAssigned        = @{ sv = 'Alla utförare'; en = 'All assignees' }
        AllTiers           = @{ sv = 'Alla supportgrupper'; en = 'All TierQueues' }
        IncidentHeading    = @{ sv = 'Incident'; en = 'Incident' }
    }

    # --- SCSM lookups ---
    $IncClass = Get-SCSMClass -Name 'System.WorkItem.Incident$'
    $RelWI = Get-SCSMRelationshipClass -Name 'System.WorkItemRelatesToWorkItem$'
    $RelAU = Get-SCSMRelationshipClass -Name 'System.WorkItemAffectedUser$'
    $RelATU = Get-SCSMRelationshipClass -Name 'System.WorkItemAssignedToUser$'
    $ActiveId = (Get-SCSMEnumeration IncidentStatusEnum.Active$).Id

    # Hämta incidenter (Status = Active)/Get incidents (Status = Active)
    $items = Get-SCSMObject -Class $IncClass -Filter "Status -eq '$ActiveId'" -ThreeLetterWindowsLanguageName $tl

    # Valfria server-side filter/Optional server-side filters
    if ($ClassificationLike) {
        $items = $items | Where-Object { $_.Classification -and $_.Classification.DisplayName -like $ClassificationLike }
    }
    if ($TierQueueLike) {
        $items = $items | Where-Object { $_.TierQueue -and $_.TierQueue.DisplayName -like $TierQueueLike }
    }

    switch ($OrderBy) {
        'CreatedDate' {
            if ($Descending) {
                $items = $items | Sort-Object -Property CreatedDate -Descending
            }
            else {
                $items = $items | Sort-Object -Property CreatedDate
            }
        }
        'Id' {
            if ($Descending) {
                $items = $items | Sort-Object -Property Id -Descending
            }
            else {
                $items = $items | Sort-Object -Property Id
            }
        }
        'Title' {
            if ($Descending) {
                $items = $items | Sort-Object -Property Title -Descending
            }
            else {
                $items = $items | Sort-Object -Property Title
            }
        }
    }

    function Get-StatusColor {
        param([string]$Status)
        switch ($Status) {
            'Active' { '#fde68a' }
            'Resolved' { '#86efac' }
            'Closed' { '#d1d5db' }
            default { '#fbbf24' }
        }
    }

    $rowsHtml = New-Object System.Text.StringBuilder
    $count = 0
    foreach ($it in $items) {
        $count++

        # relationer/relations
        $auRel = Get-SCSMRelationshipObject -BySource $it | Where-Object { $_.RelationshipId -eq $RelAU.Id } | Select-Object -First 1
        $creator = if ($auRel) { $auRel.TargetObject.DisplayName } else { '' }

        $asRel = Get-SCSMRelationshipObject -BySource $it | Where-Object { $_.RelationshipId -eq $RelATU.Id } | Select-Object -First 1
        $assigned = if ($asRel) { $asRel.TargetObject.DisplayName } else { '' }

        $rels = Get-SCSMRelationshipObject -BySource $it | Where-Object { $_.RelationshipId -eq $RelWI.Id }
        $relCount = @($rels).Count

        # fält/fields
        $id = & $htmlEncode $it.Id
        $titleEnc = & $htmlEncode $it.Title
        $creatorEnc = & $htmlEncode $creator
        $assignedEnc = & $htmlEncode $assigned
        $created = Get-Date $it.CreatedDate -Format 'yyyy-MM-dd HH:mm'
        $statusTxt = $it.Status.DisplayName
        $statusCol = Get-StatusColor $statusTxt
        $statusEnc = & $htmlEncode $statusTxt

        $classTxt = if ($it.Classification) { $it.Classification.DisplayName } else { '' }
        $classEnc = & $htmlEncode $classTxt

        $tierTxt = if ($it.TierQueue) { $it.TierQueue.DisplayName } else { '' }
        $tierEnc = & $htmlEncode $tierTxt

        $pillR = if ($relCount) { "<span class='pill'>$relCount</span>" } else { "<span class='pill pill-zero'>0</span>" }

        $details = @"
<div class='details'>
  <h5>$($t.IncidentHeading[$Language]) $id</h5>
  <table class='kv'><tbody>
    <tr><th>$($t.ID[$Language])</th><td><b>$id</b></td></tr>
    <tr><th>$($t.Title[$Language])</th><td>$titleEnc</td></tr>
    <tr><th>$($t.Status[$Language])</th><td><span class='badge' style='background-color:$statusCol'>$statusEnc</span></td></tr>
    <tr><th>$($t.Created[$Language])</th><td>$created</td></tr>
    <tr><th>$($t.AffectedUser[$Language])</th><td>$creatorEnc</td></tr>
    <tr><th>$($t.AssignedTo[$Language])</th><td>$assignedEnc</td></tr>
    <tr><th>$($t.Classification[$Language])</th><td>$classEnc</td></tr>
    <tr><th>$($t.Queue[$Language])</th><td>$tierEnc</td></tr>
  </tbody></table>
</div>
"@

        [void]$rowsHtml.AppendLine(@"
<tr class='sr-row'
    data-classification='$classEnc'
    data-AffectedUser='$creatorEnc'
    data-assignedto='$assignedEnc'
    data-tier='$tierEnc'>
  <td class='toggle-cell'><button class='toggle' aria-expanded='false' onclick='toggleRow(this)'>+</button></td>
  <td><b>$id</b></td>
  <td>$titleEnc</td>
  <td>$creatorEnc</td>
  <td>$assignedEnc</td>
  <td>$created</td>
  <td><span class='badge' style='background-color:$statusCol'>$statusEnc</span></td>
  <td>$classEnc</td>
  <td>$tierEnc</td>
  <td class='center'>$pillR</td>
</tr>
<tr class='details-row' style='display:none'>
  <td colspan='10'>$details</td>
</tr>
"@)
    }

    $generated = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    $html = @"
<!doctype html>
<html lang="$Language">
<head>
<meta charset="utf-8" />
<title>$($Title) — $count</title>
<style>
:root{--border:#E5E7EB;--bg:#F8FAFC;--bg2:#FFFFFF;--text:#111827;--muted:#6B7280;--head:#0F172A;--hover:#F1F5F9;}
*{box-sizing:border-box}
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;background:var(--bg);color:var(--text);margin:16px}
h1{margin:0 0 4px 0;color:var(--head)}
.sub{color:var(--muted);margin-bottom:12px}
.wrap{background:var(--bg2);border:1px solid var(--border);border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.toolbar{display:flex;gap:12px;align-items:center;justify-content:space-between;padding:12px 14px;border-bottom:1px solid var(--border);position:sticky;top:0;background:var(--bg2);z-index:2;border-top-left-radius:12px;border-top-right-radius:12px}
.controls{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.btn{padding:6px 10px;border:1px solid var(--border);background:#fff;border-radius:8px;cursor:pointer}
.btn:hover{background:var(--hover)}
.select{padding:6px 8px;border:1px solid var(--border);border-radius:8px;min-width:200px;background:#fff}
table.master{border-collapse:separate;border-spacing:0;width:100%;table-layout:fixed}
table.master th,table.master td{border-bottom:1px solid var(--border);padding:10px;text-align:left;vertical-align:middle}
table.master th:first-child,table.master td:first-child{width:44px;text-align:center}
thead th{position:sticky;top:58px;background:linear-gradient(180deg,#F1F5F9,#E5E7EB);color:#0f172a;font-weight:600;z-index:1}
td.center{text-align:center}
.pill{display:inline-block;min-width:24px;padding:2px 8px;border-radius:999px;border:1px solid var(--border);background:#F3F4F6;text-align:center;font-weight:600}
.pill-zero{opacity:.6}
.badge{display:inline-block;padding:3px 9px;border-radius:999px;border:1px solid rgba(0,0,0,0.05);font-weight:600}
.details{padding:12px;background:#F8FAFF;border:1px solid #DBEAFE;border-radius:10px}
.toggle{width:30px;height:30px;border:1px solid var(--border);background:#fff;border-radius:8px;cursor:pointer;font-weight:700}
.toggle[aria-expanded="true"]{background:#DBEAFE;border-color:#BFDBFE;color:#1E3A8A}
</style>
<script>
var clientLang = '$Language';
function toggleRow(btn){
  var detailRow=btn.closest('tr').nextElementSibling;if(!detailRow)return;
  var expanded=btn.getAttribute("aria-expanded")==="true";
  detailRow.style.display=expanded?"none":"table-row";
  btn.setAttribute("aria-expanded",(!expanded).toString());
  btn.textContent=expanded?"+":"–";
}
function expandAll(){document.querySelectorAll("tr.details-row").forEach(r=>r.style.display="table-row");
  document.querySelectorAll("button.toggle").forEach(b=>{b.setAttribute("aria-expanded","true");b.textContent="–";});}
function collapseAll(){document.querySelectorAll("tr.details-row").forEach(r=>r.style.display="none");
  document.querySelectorAll("button.toggle").forEach(b=>{b.setAttribute("aria-expanded","false");b.textContent="+";});}
function norm(s){return (s||"").toLowerCase().trim();}
function populateSelect(id,placeholder,set){
  var el=document.getElementById(id); if(!el) return;
  el.innerHTML=""; var o=document.createElement("option");o.value="";o.textContent=placeholder;el.appendChild(o);
  Array.from(set).sort((a,b)=>a.localeCompare(b, clientLang)).forEach(v=>{
    var x=document.createElement("option"); x.value=v; x.textContent=v; el.appendChild(x);
  });
}
function applyFilters(){
  var vClass=norm(document.getElementById('classSelect').value);
  var vCreator=norm(document.getElementById('AffectedUserSelect').value);
  var vAss=norm(document.getElementById('assignedToSelect').value);
  var vTier=norm(document.getElementById('tierSelect').value);
  document.querySelectorAll('table.master tbody tr.sr-row').forEach(function(row){
    var c=norm(row.getAttribute('data-classification'));
    var cu=norm(row.getAttribute('data-AffectedUser'));
    var as=norm(row.getAttribute('data-assignedto'));
    var t=norm(row.getAttribute('data-tier'));
    var show=(vClass===''||c===vClass)&&(vCreator===''||cu===vCreator)&&(vAss===''||as===vAss)&&(vTier===''||t===vTier);
    row.style.display=show?'':'none';
    var d=row.nextElementSibling; if(d && d.classList.contains('details-row') && !show){
      d.style.display='none'; var b=row.querySelector('button.toggle'); if(b){b.setAttribute('aria-expanded','false');b.textContent='+';}
    }
  });
}
window.addEventListener('DOMContentLoaded', function(){
  var classSet=new Set(), creatorSet=new Set(), assignedSet=new Set(), tierSet=new Set();
  document.querySelectorAll('tbody tr.sr-row').forEach(function(r){
    var cl=r.getAttribute('data-classification')||''; if(cl.trim().length) classSet.add(cl);
    var cu=r.getAttribute('data-AffectedUser')||'';      if(cu.trim().length) creatorSet.add(cu);
    var as=r.getAttribute('data-assignedto')||'';     if(as.trim().length) assignedSet.add(as);
    var ti=r.getAttribute('data-tier')||'';           if(ti.trim().length) tierSet.add(ti);
  });
  populateSelect('classSelect','${($t.AllClassifications[$Language])}',classSet);
  populateSelect('AffectedUserSelect','${($t.AllCreators[$Language])}',creatorSet);
  populateSelect('assignedToSelect','${($t.AllAssigned[$Language])}',assignedSet);
  populateSelect('tierSelect','${($t.AllTiers[$Language])}',tierSet);
  ['classSelect','AffectedUserSelect','assignedToSelect','tierSelect'].forEach(function(id){
    var el=document.getElementById(id); if(el){ el.addEventListener('change',applyFilters); }
  });
});
</script>
</head>
<body>
  <div class="wrap">
    <div class="toolbar">
      <div>
        <h1>$($Title)</h1>
        <div class="sub">$([string]::Format($t.IncidentsCount[$Language], $count, $generated))</div>
      </div>
      <div class="controls">
        <select id="classSelect" class="select"></select>
        <select id="AffectedUserSelect" class="select"></select>
        <select id="assignedToSelect" class="select"></select>
        <select id="tierSelect" class="select"></select>
        <button class="btn" onclick="expandAll()">$($t.ExpandAll[$Language])</button>
        <button class="btn" onclick="collapseAll()">$($t.CollapseAll[$Language])</button>
      </div>
    </div>

    <table class="master">
      <thead>
        <tr>
          <th></th>
          <th>$($t.ID[$Language])</th>
          <th>$($t.Title[$Language])</th>
          <th>$($t.AffectedUser[$Language])</th>
          <th>$($t.AssignedTo[$Language])</th>
          <th>$($t.Created[$Language])</th>
          <th>$($t.Status[$Language])</th>
          <th>$($t.Classification[$Language])</th>
          <th>$($t.Queue[$Language])</th>
          <th>$($t.Related[$Language])</th>
        </tr>
      </thead>
      <tbody>
        $($rowsHtml.ToString())
      </tbody>
    </table>
  </div>
</body>
</html>
"@

    $dir = Split-Path -Path $OutputPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
    Get-Item $OutputPath
}
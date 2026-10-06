<#
.SYNOPSIS
    Hugo Stack 博客文章管理脚本：新建 / 列表 / 删除文章（支持中英双语）
.EXAMPLE
    # 双击 manage-blog.bat，或在项目目录运行：
    powershell -ExecutionPolicy Bypass -File .\manage-posts.ps1
.EXAMPLE
    # 非交互式（可直接传参调用）
    .\manage-posts.ps1 -List
    .\manage-posts.ps1 -New -Title "我的第一篇" -EnTitle "My First Post" -Categories @("随笔") -Tags @("Hugo")
    .\manage-posts.ps1 -RemoveSlug "my-first-post" -Force
#>

[CmdletBinding(DefaultParameterSetName = "Menu")]
param(
    [Parameter(ParameterSetName = "List")]
    [switch]$List,

    [Parameter(ParameterSetName = "New")]
    [switch]$New,

    [Parameter(ParameterSetName = "New", Mandatory = $true)]
    [string]$Title,

    [Parameter(ParameterSetName = "New")]
    [string]$EnTitle = "",

    [Parameter(ParameterSetName = "New")]
    [string]$Slug = "",

    [Parameter(ParameterSetName = "New")]
    [string]$Description = "",

    [Parameter(ParameterSetName = "New")]
    [string]$EnDescription = "",

    [Parameter(ParameterSetName = "New")]
    [string[]]$Categories = @(),

    [Parameter(ParameterSetName = "New")]
    [string[]]$Tags = @(),

    [Parameter(ParameterSetName = "New")]
    [switch]$Draft,

    [Parameter(ParameterSetName = "RemoveSlug")]
    [switch]$Remove,

    [Parameter(ParameterSetName = "RemoveSlug", Mandatory = $true)]
    [string]$RemoveSlug,

    [Parameter(ParameterSetName = "RemoveSlug")]
    [switch]$Force
)

# --- 控制台 UTF-8，保证中文正常显示 ---
chcp 65001 > $null
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$PostRoot    = Join-Path $ProjectRoot "content\post"

function Read-FrontMatter {
    param([string]$Path)
    $meta = [ordered]@{ Title = ""; Date = ""; Draft = $false }
    if (-not (Test-Path $Path)) { return $meta }
    $lines = Get-Content -LiteralPath $Path -Encoding UTF8
    foreach ($line in $lines) {
        if ($line -match '^\s*title:\s*"?(.*?)"?\s*$' -and -not $meta.Title) {
            $meta.Title = $Matches[1].Trim()
        }
        elseif ($line -match '^\s*date:\s*["'']?(\d{4}-\d{2}-\d{2})') {
            $meta.Date = $Matches[1]
        }
        elseif ($line -match '^\s*draft:\s*true') {
            $meta.Draft = $true
        }
    }
    return $meta
}

function Get-BlogPosts {
    $items = @()
    if (-not (Test-Path $PostRoot)) { return $items }
    Get-ChildItem -LiteralPath $PostRoot -Directory | Sort-Object Name | ForEach-Object {
        $dir   = $_
        $zh    = Join-Path $dir.FullName "index.zh.md"
        $en    = Join-Path $dir.FullName "index.en.md"
        $meta  = $null
        if (Test-Path $zh)       { $meta = Read-FrontMatter $zh }
        elseif (Test-Path $en)   { $meta = Read-FrontMatter $en }
        if ($null -ne $meta) {
            $items += [pscustomobject]@{
                Slug      = $dir.Name
                Title     = $(if ($meta.Title) { $meta.Title } else { $dir.Name })
                Date      = $meta.Date
                Draft     = $meta.Draft
                Chinese   = Test-Path $zh
                English   = Test-Path $en
                Directory = $dir.FullName
            }
        }
    }
    return $items
}

function Show-Posts {
    $items = Get-BlogPosts
    if ($items.Count -eq 0) {
        Write-Host "  （当前还没有文章）" -ForegroundColor DarkGray
        return $items
    }
    $i = 0
    foreach ($p in $items) {
        $i++
        $langs = @()
        if ($p.Chinese) { $langs += "中" }
        if ($p.English) { $langs += "EN" }
        $tag = "[" + ($langs -join "/") + "]"
        $draftMark = if ($p.Draft) { "  [草稿]" } else { "" }
        "{0,3}. {1} {2,-12} {3}{4}" -f $i, $tag, $p.Date, $p.Title, $draftMark | Write-Host
    }
    return $items
}

function Convert-ToSlug {
    param([string]$Text)
    # 去掉 Windows 文件名非法字符与空白，保留中英文、数字、连字符、下划线
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ""
    $pattern = [Regex]::Escape($invalid)
    $clean = [Regex]::Replace($Text, "[$pattern]", "")
    $clean = $clean -replace '\s+', '-'
    return $clean.Trim('-')
}

function Build-Content {
    param(
        [string]$PostTitle,
        [string]$PostDescription,
        [string]$PostDate,
        [string[]]$PostCategories,
        [string[]]$PostTags,
        [bool]$IsDraft
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("---")
    [void]$sb.AppendLine("title: `"$PostTitle`"")
    [void]$sb.AppendLine("description: `"$PostDescription`"")
    [void]$sb.AppendLine("date: $PostDate")
    if ($PostCategories.Count -gt 0) {
        [void]$sb.AppendLine("categories:")
        foreach ($c in $PostCategories) { [void]$sb.AppendLine("    - $c") }
    }
    if ($PostTags.Count -gt 0) {
        [void]$sb.AppendLine("tags:")
        foreach ($t in $PostTags) { [void]$sb.AppendLine("    - $t") }
    }
    [void]$sb.AppendLine('image: ""')
    [void]$sb.AppendLine("comments: true")
    [void]$sb.AppendLine("draft: $($IsDraft.ToString().ToLower())")
    [void]$sb.AppendLine("---")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("<!-- 在这里开始写作，支持 Markdown 语法 -->")
    [void]$sb.AppendLine("")
    return $sb.ToString()
}

function New-BlogPost {
    param(
        [string]$PostTitle,
        [string]$PostEnTitle,
        [string]$PostSlug,
        [string]$PostDescription,
        [string]$PostEnDescription,
        [string[]]$PostCategories,
        [string[]]$PostTags,
        [bool]$IsDraft
    )

    if ([string]::IsNullOrWhiteSpace($PostTitle)) {
        throw "文章标题不能为空"
    }

    # 生成目录名（slug）
    if ([string]::IsNullOrWhiteSpace($PostSlug)) {
        $PostSlug = Get-Date -Format "yyyyMMdd-HHmmss"
    }
    else {
        $PostSlug = Convert-ToSlug $PostSlug
    }
    if (-not $PostSlug) { throw "目录名不合法，请重新输入" }

    $target = Join-Path $PostRoot $PostSlug
    if (Test-Path $target) {
        throw "目录已存在：content\post\$PostSlug，请换一个名称"
    }

    $date = Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz"
    New-Item -ItemType Directory -Path $target -Force | Out-Null

    $zhBody = Build-Content -PostTitle $PostTitle -PostDescription $PostDescription -PostDate $date `
        -PostCategories $PostCategories -PostTags $PostTags -IsDraft $IsDraft
    [IO.File]::WriteAllText((Join-Path $target "index.zh.md"), $zhBody, (New-Object System.Text.UTF8Encoding($false)))

    $created = @("content\post\$PostSlug\index.zh.md")

    if (-not [string]::IsNullOrWhiteSpace($PostEnTitle)) {
        $enBody = Build-Content -PostTitle $PostEnTitle -PostDescription $PostEnDescription -PostDate $date `
            -PostCategories $PostCategories -PostTags $PostTags -IsDraft $IsDraft
        [IO.File]::WriteAllText((Join-Path $target "index.en.md"), $enBody, (New-Object System.Text.UTF8Encoding($false)))
        $created += "content\post\$PostSlug\index.en.md"
    }

    return [pscustomobject]@{ Slug = $PostSlug; Files = $created; Directory = $target }
}

function Split-Indexes {
    param([string]$InputText, [int]$Max)
    $result = @()
    foreach ($part in ($InputText -split '[,，\s]+')) {
        $n = 0
        if ([int]::TryParse($part.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $Max) {
            if ($result -notcontains $n) { $result += $n }
        }
    }
    return $result
}

function Remove-BlogPosts {
    param([int[]]$Indexes)
    $items = Get-BlogPosts
    if ($items.Count -eq 0) { Write-Host "  没有可删除的文章。" -ForegroundColor DarkGray; return }

    Write-Host ""
    Write-Host "即将删除以下文章（不可恢复）：" -ForegroundColor Yellow
    foreach ($idx in $Indexes) {
        $p = $items[$idx - 1]
        Write-Host ("  - [{0}] {1}  (content\post\{2})" -f $idx, $p.Title, $p.Slug)
    }
    $answer = Read-Host "确认删除？输入 y 确认，其他键取消"
    if ($answer -ne 'y' -and $answer -ne 'Y') {
        Write-Host "已取消。" -ForegroundColor DarkGray
        return
    }
    foreach ($idx in $Indexes) {
        $p = $items[$idx - 1]
        Remove-Item -LiteralPath $p.Directory -Recurse -Force
        Write-Host ("已删除：{0}" -f $p.Title) -ForegroundColor Green
    }
    Write-Host "提示：若本地预览仍能打开被删文章，重启 hugo server 即可生效。" -ForegroundColor DarkGray
}

# ============ 非交互式（参数）入口 ============
if ($PSCmdlet.ParameterSetName -eq "List") {
    $null = Show-Posts
    return
}

if ($PSCmdlet.ParameterSetName -eq "New") {
    # 兼容命令行传入 "分类A,分类B" 这种逗号字符串
    $Categories = @($Categories | ForEach-Object { $_ -split '[,，]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $Tags       = @($Tags       | ForEach-Object { $_ -split '[,，]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $r = New-BlogPost -PostTitle $Title -PostEnTitle $EnTitle -PostSlug $Slug `
        -PostDescription $Description -PostEnDescription $EnDescription `
        -PostCategories $Categories -PostTags $Tags -IsDraft ([bool]$Draft)
    Write-Host ("文章已创建：{0}" -f $r.Slug) -ForegroundColor Green
    $r.Files | ForEach-Object { Write-Host "  $_" }
    return
}

if ($PSCmdlet.ParameterSetName -eq "RemoveSlug") {
    $target = Join-Path $PostRoot $RemoveSlug
    if (-not (Test-Path $target)) { throw "文章不存在：content\post\$RemoveSlug" }
    if (-not $Force) {
        $answer = Read-Host "确认删除 content\post\$RemoveSlug ？输入 y 确认"
        if ($answer -ne 'y' -and $answer -ne 'Y') { Write-Host "已取消。"; return }
    }
    Remove-Item -LiteralPath $target -Recurse -Force
    Write-Host "已删除：content\post\$RemoveSlug" -ForegroundColor Green
    Write-Host "提示：若本地预览仍能打开被删文章，重启 hugo server 即可生效。" -ForegroundColor DarkGray
    return
}

# ============ 交互式菜单 ============
while ($true) {
    Clear-Host
    Write-Host ""
    Write-Host "  ==========================================" -ForegroundColor Cyan
    Write-Host "        Hugo 博客文章管理" -ForegroundColor Cyan
    Write-Host "  ==========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    1. 新建文章"
    Write-Host "    2. 删除文章"
    Write-Host "    3. 查看全部文章"
    Write-Host "    0. 退出"
    Write-Host ""
    $choice = Read-Host "请输入选项 (0-3)"

    switch ($choice) {
        "1" {
            Clear-Host
            Write-Host "--- 新建文章 ---" -ForegroundColor Cyan
            try {
                $t  = Read-Host "中文标题（必填）"
                $et = Read-Host "英文标题（留空则只建中文版，回车跳过）"
                $d  = Read-Host "文章摘要（可留空）"
                $ed = Read-Host "英文摘要（可留空）"
                $c  = Read-Host "分类（多个用逗号分隔，可留空）"
                $g  = Read-Host "标签（多个用逗号分隔，可留空）"
                $s  = Read-Host "目录名 slug（留空自动用日期时间，如 20261007-103015）"
                $df = Read-Host "是否存为草稿？发布后不在首页显示 (y/N)"

                $cats = @($c -split '[,，]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                $tags = @($g -split '[,，]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                $isDraft = ($df -eq 'y' -or $df -eq 'Y')

                $r = New-BlogPost -PostTitle $t -PostEnTitle $et -PostSlug $s `
                    -PostDescription $d -PostEnDescription $ed `
                    -PostCategories $cats -PostTags $tags -IsDraft $isDraft

                Write-Host ""
                Write-Host ("文章已创建：{0}" -f $r.Slug) -ForegroundColor Green
                $r.Files | ForEach-Object { Write-Host "  $_" }
                $open = Read-Host "是否打开文章所在文件夹？(y/N)"
                if ($open -eq 'y' -or $open -eq 'Y') { Invoke-Item $r.Directory }
            }
            catch {
                Write-Host ""
                Write-Host ("创建失败：{0}" -f $_.Exception.Message) -ForegroundColor Red
            }
            Write-Host ""
            Read-Host "按回车返回菜单" | Out-Null
        }
        "2" {
            Clear-Host
            Write-Host "--- 删除文章 ---" -ForegroundColor Cyan
            $items = @(Get-BlogPosts)
            $null = Show-Posts
            if ($items.Count -gt 0) {
                $raw = Read-Host "输入要删除的序号（多个用逗号分隔，如 1,3；回车取消）"
                $idx = @(Split-Indexes -InputText $raw -Max $items.Count)
                if ($idx.Count -gt 0) {
                    Remove-BlogPosts -Indexes $idx
                } else {
                    Write-Host "未选择有效序号，已取消。" -ForegroundColor DarkGray
                }
            }
            Write-Host ""
            Read-Host "按回车返回菜单" | Out-Null
        }
        "3" {
            Clear-Host
            Write-Host "--- 全部文章 ---" -ForegroundColor Cyan
            Write-Host ""
            $null = Show-Posts
            Write-Host ""
            Read-Host "按回车返回菜单" | Out-Null
        }
        "0" { break }
        default {
            Write-Host "无效选项。" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
    if ($choice -eq "0") { break }
}


class Operation {
    [string]$Name
    [scriptblock]$Script
    
    Operation([string]$Name, [scriptblock]$Script) {
        $this.Name = $Name
        $this.Script = $Script
    }
    
    # 执行操作
    [void]Execute() {
        try {
            & $this.Script
        }
        catch {
            Write-Host "执行操作失败: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

class CopyToVmParam {
    [string]$Path
    [string]$Destination
    [bool]$Recurse = $false
    [bool]$Force = $false
    [bool]$Verbose = $false
    

    [void]CopyToVm([string]$VhdDriveLetter) {
        $safeDriveLetter = ([string]$VhdDriveLetter).Trim()
        # 替换盘符 c:xxxx > g:xxxx
        $finalDestination = $this.Destination -replace '^[\w]+(?=:)', $safeDriveLetter

        if (-not (Test-Path -Path $finalDestination)) {
            New-Item -ItemType Directory -Path $finalDestination
        }
        
        Copy-Item -Path $this.Path -Destination $finalDestination -Recurse:$this.Recurse -Force:$this.Force -Verbose:$this.Verbose
    }
}

class GpuInfo {
    # Get-PnpDevice
    [string]$FriendlyName
    [string]$ServiceName
    # Get-WmiObject Win32_PNPSignedDriver
    [string]$DeviceId
    [string]$DeviceName
    [string]$DeviceManufacturer
    # Get-VMHostPartitionableGpu
    [string]$InstancePath
    # regdit
    [Int64]$MemorySize
    
    [string]$AdapterId
    [string]$VMName
}

function Write-HostColored {
    param (
        [string]$content,
        [switch]$NoNewline
    )

    # 定义允许的颜色列表
    $allowedColors = @('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')
    $defaultColor = "Blue"  # 默认颜色为蓝色

    # 定义正则表达式，匹配 {{@Color@Content}} 或 {{Content}} 的部分
    $regex = '{{ @(\w+?)@(.*?) }}|{{ (.*?) }}'

    # 用正则表达式查找所有匹配 {{@Color@Content}} 或 {{Content}} 的部分
    $matches = [regex]::Matches($content, $regex)

    # 初始化位置
    $currentPosition = 0
    
    # 遍历每一个匹配到的部分
    foreach ($match in $matches) {
        # 打印匹配到的普通文本部分
        $startIndex = $match.Index
        $lengthBeforeHighlight = $startIndex - $currentPosition
        $beforeHighlight = $content.Substring($currentPosition, $lengthBeforeHighlight)
        Write-Host -NoNewline $beforeHighlight

        # 初始化颜色和文本
        $color = $defaultColor
        $text = $match.Groups[3].Value

        # 如果文本为空，说明匹配的是 {{@Color@Content}}
        if ($text -eq "") {
            $color = $match.Groups[1].Value
            $text = $match.Groups[2].Value
            # 颜色不匹配时使用默认颜色
            if (-not ($allowedColors -contains $color)) {
                $color = $defaultColor
            }
        }

        # 动态设置颜色并打印内容
        Write-Host -NoNewline -ForegroundColor $color $text

        # 更新当前位置
        $currentPosition = $match.Index + $match.Length
    }

    # 打印最后的普通文本部分
    $lastText = ""
    if ($currentPosition -lt $content.Length) {
        $lastText = $content.Substring($currentPosition)
    }
    if ($NoNewline) {
        Write-Host -NoNewline $lastText
    }
    else {
        Write-Host $lastText
    }
}

function Get-UserInputWithOptions {
    param (
        [string]$Prompt,
        [switch]$Number = $false, 
        [switch]$Integer = $false,
        [string]$Match,
        [double]$LT = [double]::NaN,
        [double]$GT = [double]::NaN,
        [double]$LE = [double]::NaN,
        [double]$GE = [double]::NaN,
        [string[]]$Options = @()
    )
    
    if (-not $Number) {
        $LT, $GT, $LE, $GE | ForEach-Object {
            if ($_ -gt [double]::NaN) {
                $Number = $true
            }
        }
    }

    while ($true) {
        
        while ($true) {
            Write-HostColored -NoNewline $Prompt
            $userInput = Read-Host

            # 校验是否在可选项列表中
            if ($Options.Count -gt 0) {
                if ($Options -notcontains $userInput) {
                    Write-Host "输入无效: 请输入以下选项之一: $($Options -join ', ')" -ForegroundColor Red
                    continue
                }
            }

            if ($Match -ne "" -and $userInput -notmatch $Match) {
                Write-Host "输入无效" -ForegroundColor Red
            }

            # 数字校验
            if ($Number -and $userInput -notmatch '^-?\d+$' ) {
                Write-Host "输入必须是一个有效的数字" -ForegroundColor Red
                continue
            }

            if ($Integer -and $userInput -notmatch '^-\d+$') {
                Write-Host "输入必须是一个整数" -ForegroundColor Red
            }

            if ($Number) {
                $longInput = [long]$userInput

                if ($LT -gt [long]::MinValue -and -not($longInput -lt $LT)) {
                    Write-Host "输入必须小于 $LT" -ForegroundColor Red
                    continue
                }

                if ($GT -gt [long]::MinValue -and -not($longInput -gt $GT)) {
                    Write-Host "输入必须大于 $GT" -ForegroundColor Red
                    continue
                }

                if ($LE -gt [long]::MinValue -and -not($longInput -le $LE)) {
                    Write-Host "输入必须小于或的等于 $LE" -ForegroundColor Red
                    continue
                }

                if ($GE -gt [long]::MinValue -and -not($longInput -ge $GE)) {
                    Write-Host "输入必须大于或等于 $GE" -ForegroundColor Red
                    continue
                }
            }
            
            return $userInput
        }
    }
}

function Select-Item() {
    param(
        [array]$Items,
        [array]$Properties = @(),
        [string]$ItemDescription = "选项"
    )

    $no = 1
    $showItems = @()
    $dictNoToShowItems = @{}
    $dictNoToItem = @{}
    $Items | ForEach-Object {
        $showItem = [PSCustomObject]@{
            '#' = [string]$no
        }
        foreach ($prop in $_.psobject.properties) {
            if ($Properties.Length -eq 0 -or $Properties.Contains($prop.Name)) {
                $showItem | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
            }
        }

        $dictNoToItem[[string]$no] = $_
        $dictNoToShowItems[[string]$no] = $showItem
        $showItems += $showItem
        $no++
    }

    Write-HostColored -NoNewline "{{ $ItemDescription }}清单："
    $showItems | Format-Table -Wrap -AutoSize | Out-Host

    $selected = $null
    while (-not $selected) {
        Write-HostColored -NoNewline "输入序号选择对应的{{ $ItemDescription }}："
        $choice = Read-Host

        if ($dictNoToItem.ContainsKey($choice)) {
            $selected = $dictNoToItem[$choice]
        }
        else {
            Write-Host "无效输入: $choice" -ForegroundColor Red
        }
    }

    $dictNoToShowItems[$choice] | Format-List | Out-Host
    return $selected
}

function ConvertTo-HaedwareId {
    param (
        [string]$InstancePath
    )
    return "PCI\$($_.Name.Split('#')[1])"
}

function Get-HostGpu {
    param(
        [string]$InstancePath
    )

    $infos = Get-VMHostPartitionableGpu | Where-Object -FilterScript {
        -not $InstancePath -or $_.Name -eq $InstancePath
    } | ForEach-Object {
        $hardwareID = ConvertTo-HaedwareId $_.Name
        $pnpSignedDevice = Get-WmiObject Win32_PNPSignedDriver | Where-Object -Property HardwareID -EQ $hardwareID
        $pnpDevice = Get-PnpDevice -Class "Display" -DeviceId $pnpSignedDevice.DeviceId

        # 获取显存
        $value = Get-ItemProperty -Path "HKLM:HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\$($pnpSignedDevice.DeviceId)"
        $value = Get-ItemProperty -Path "HKLM:HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\$($value.Driver)"
        $memorySize = $value."HardwareInformation.qwMemorySize"
    
        $info = [GpuInfo]::new()
        $info.FriendlyName = $pnpDevice.FriendlyName
        $info.ServiceName = $pnpDevice.Service
    
        $info.DeviceId = $pnpSignedDevice.DeviceId
        $info.DeviceName = $pnpSignedDevice.DeviceName
        $info.DeviceManufacturer = $pnpSignedDevice.Manufacturer
        $info.InstancePath = $_.Name
        $info.MemorySize = $memorySize
        $info
    }
    return $infos
}

function Select-HostGpu {
    $infos = Get-HostGpu
    $info = Select-Item -Items $infos -ItemDescription "Gpu" -Properties FriendlyName, DeviceManufacturer
    return $info
}

function Select-VM {
    $vms = Get-VM
    $vm = Select-Item -Items $vms -ItemDescription "VM" -Properties Name, State, Status, MemoryAssigned, MemoryStartup, DynamicMemoryEnabled

    if ($vm.State -eq 'Running') {
        Write-Host "VM($($vm.Name)) 正在运行，请关闭后再操作。" -ForegroundColor Red
        Exit -1
    }

    return $vm
}

function Get-VMGpu {
    param (
        [string]$VMName
    )

    $dictInstanceToHostGpu = @{}
    $infos = Get-VMGpuPartitionAdapter -VMName $VMName | ForEach-Object {
        $info = $dictInstanceToHostGpu[$_.InstancePath]
        if ($null -eq $info) {
            $info = Get-HostGpu -InstancePath $_.InstancePath
            $dictInstanceToHostGpu[$_.InstancePath] = $info
        }
        else {
            $info = $info | Select-Object -Property *
        }
        $info.VMName = $VMName
        $info.AdapterId = $_.Id
        $info
    }
    return $infos
}

function Show-VMGpu {
    param (
        [string]$VMName
    )
    $vmGpuInfo = Get-VMGpu -VMName $VMName
    if ($null -eq $vmGpuInfo) {
        Write-HostColored "VM({{ $($vm.Name) }})没有分配任何Gpu"
    }
    else {
        Write-HostColored "VM({{ $($vm.Name) }})Gpu分配详情:"
        $vmGpuInfo | Format-Table -Property FriendlyName, AdapterId -Wrap -AutoSize | Out-Host
    }
}

function Select-VMGpu {
    param(
        [string]$VMName
    )

    if ($VMName -eq "") {
        $vm = Select-VM
        $VMName = $vm.Name
    }

    $infos = Get-VMGpu -VMName $VMName

    $info = Select-Item -Items $infos -ItemDescription "VM($VMName) 已分配 Gpu" -Properties FriendlyName, AdapterId
    return $info
}

function Set-GpuToVm {
    $vm = Select-VM
    $gpu = Select-HostGpu

    Write-HostColored "为虚拟机进行核心资源分配："
    Write-HostColored "Cpu核心：{{ 4 }}"
    Write-HostColored "检查点功能：{{ 禁用 }}"
    Write-HostColored "低内存映射 I/O 空间：{{ 3GB }}"
    Write-HostColored "高内存映射 I/O 空间：{{ 32GB }}"
    Write-HostColored "允许虚拟机操作系统控制缓存类型：{{ 开启 }}"
    Write-HostColored "主机关闭或重启时自动关闭虚拟机：{{ 开启 }}"

    Set-VM -Name $vm.Name `
        -ProcessorCount 4 `
        -CheckpointType Disabled `
        -LowMemoryMappedIoSpace 3GB -HighMemoryMappedIoSpace 32GB `
        -GuestControlledCacheTypes $true `
        -AutomaticStopAction ShutDown
    
    Add-VMGpuPartitionAdapter -VMName $vm.Name -InstancePath $gpu.InstancePath
    Set-VMGpuPartitionAdapter -VMName $vm.Name

    Update-VMGpuDriver -VMName $vm.Name -GpuInfo $gpu

    Write-HostColored "为 VM({{ $($vm.Name) }}) 分配 Gpu({{ $($gpu.FriendlyName) }}) 成功"

    Show-VMGpu -VMName $vm.Name
}

function Remove-GpuFromVm {
    $vm = Select-VM
    
    $gpu = Select-VMGpu -VMName $vm.Name

    Remove-VMGpuPartitionAdapter -VMName $vm.Name -AdapterId $gpu.AdapterId
    Write-HostColored "从 VM({{ $($vm.Name) }}) 移除 Gpu({{ $($gpu.FriendlyName) }}) 成功"

    Show-VMGpu -VMName $vm.Name
}

function Copy-ToVmSystemVolume() {
    param (
        [string]$VMName,
        [array]$CopyParams
    )

    $vmSystemHd = Get-VMHardDiskDrive -VMName $VMName | Select-Object -First 1
    if (-not $vmSystemHd) {
        throw "未找到虚拟机系统硬盘驱动。"
    }

    # 获取 VHD 信息
    $vhd = Get-VHD $vmSystemHd.Path
    $attached = $vhd.Attached

    # 挂载 VHD 如果未挂载
    if (-not $attached) {
        Mount-VHD -Path $vmSystemHd.Path
    }

    # 获取已挂载的 VHD 的信息
    $vhdInfo = Get-Disk | Where-Object { $_.Location -eq $vmSystemHd.Path }

    # 获取卷的盘符
    $volume = Get-Partition -DiskNumber $vhdInfo.Number | Get-Volume
    
    # 复制文件到 VM
    Write-Host "开始将文件复制到虚拟机..."
    $CopyParams | ForEach-Object {
        $_.CopyToVm($volume.DriveLetter)
    }
    Write-Host "文件复制完成。"

    if (-not $attached) {
        Dismount-VHD -Path $vmSystemHd.Path
    }

}

function Update-VMGpuDriver {
    param (
        [string]$VMName,
        [GpuInfo]$GpuInfo
    )

    if ($VMName -eq "") {
        $vm = Select-VM
        $VMName = $vm.Name
    }
    if ($null -eq $GpuInfo) {
        $GpuInfo = Select-VMGpu -VMName $VMName
    }

    Write-Host "搜索主机驱动文件，需要几分钟，请耐心等待"

    $serviceDrivce = Get-WmiObject Win32_SystemDriver | Where-Object -Property Name -EQ $GpuInfo.ServiceName
    $driverFiles = @()
    $driverFiles += $serviceDrivce.PathName

    $modifiedDeviceId = $GpuInfo.DeviceId.Replace("\", "\\")
    $Antecedent = "\\$ENV:COMPUTERNAME\ROOT\cimv2:Win32_PNPSignedDriver.DeviceID=""$modifiedDeviceId"""
    Get-WmiObject Win32_PNPSignedDriverCIMDataFile | Where-Object -Property Antecedent -EQ $Antecedent | ForEach-Object {
        # \\MARIOPLUSMAIN\ROOT\cimv2:CIM_DataFile.Name="c:\\windows\\inf\\oem27.inf"
        # c:\\windows\\inf\\oem27.inf
        $path = $_.Dependent.Split("=")[1].Trim('"')
    
        # c:\windows\inf\oem27.inf
        $normalizePath = [System.IO.Path]::GetFullPath($path)
        $driverFiles += $normalizePath
    }
    $copyParams = $driverFiles | ForEach-Object {
        $vmPath = $_ -replace '(?i)(c:\\windows\\system32\\)driverstore\\', '${1}HostDriverStore\'
    
        $copyParam = [CopyToVmParam]::new()
        $copyParam.Path = $_
        $copyParam.Destination = [System.IO.Path]::GetDirectoryName($vmPath)
        $copyParam.Recurse = $true
        $copyParam
    }

    Copy-ToVmSystemVolume -VMName $VMName -CopyParams $copyParams
}

function Set-PartitionResources {
    param (
        [string]$VMName,
        [float]$BaseValue, # 资源的基准值
        [string]$ResourceType, # 资源类型 (VRAM, Encode, Decode, Compute)
        [int]$Percentage = 100
    )
    # 计算比例除数
    [float]$devider = [math]::round(100 / $Percentage, 2)

    $minValue = [math]::round($BaseValue / $devider)
    $maxValue = $minValue
    $optimalValue = $minValue

    # 使用动态参数名设置对应资源
    Write-HostColored "设置{{ $ResourceType }}资源, Mix: {{ $minValue }}, Max: {{ $maxValue }}, Optimal:{{ $optimalValue }}"
    $params = @{
        VMName                          = $VMName
        "MinPartition$ResourceType"     = $minValue
        "MaxPartition$ResourceType"     = $maxValue
        "OptimalPartition$ResourceType" = $optimalValue
    }
    Set-VMGpuPartitionAdapter @params
}

function Set-VMGpuResource() {
    param(
        [string]$VMName,
        [GpuInfo]$GpuInfo,
        [int]$Percentage
    )

    if ($VMName -eq "") {
        $vm = Select-VM
        $VMName = $vm.Name
    }

    if ($null -eq $GpuInfo) {
        $GpuInfo = Select-VMGpu -VMName $VMName
    }

    while ($null -eq $Percentage -or $Percentage -lt 1 -or $Percentage -gt 100) {
        Write-HostColored -NoNewline "请输入为VM({{ $VMName }})Gpu({{ $($GpuInfo.FriendlyName) }})分配的资源占比(1-100):"
        $r = Read-Host
        if ($r -match "^\d+$") {
            $Percentage = $r
        }
    }
    
    # 定义基准值
    $baseVRAM = 1000000000 # 显存基准值 1GB
    $baseEncode = 18446744073709551615 # 编码资源最大值
    $baseDecode = 1000000000 # 解码资源基准值 1GB
    $baseCompute = 1000000000 # 计算资源基准值 1GB

    # 设置显存
    Set-PartitionResources -VMName $VMName -BaseValue $baseVRAM -resourceType "VRAM" -Percentage $Percentage

    # 设置编码资源
    Set-PartitionResources -VMName $VMName -BaseValue $baseEncode -resourceType "Encode" -Percentage $Percentage

    # 设置解码资源
    Set-PartitionResources -VMName $VMName -BaseValue $baseDecode -resourceType "Decode" -Percentage $Percentage

    # 设置计算资源
    Set-PartitionResources -VMName $VMName -BaseValue $baseCompute -resourceType "Compute" -Percentage $Percentage


}

function Show-VM() {
    param(
        [string]$VMName
    )

    if ($VMName -eq "") {
        $infos = Get-VM
    }
    else {
        $infos = Get-VM -Name $VMName
    }

    $no = 1
    $infos | ForEach-Object {
        $info = [PSCustomObject]@{
            "#"     = $no
            "ID"    = $_.Id
            "名称"    = $_.Name
            "状态"    = $_.State
            "已分配内存" = "$($_.MemoryAssigned / 1GB)GB"
            "动态内存"  = $_.DynamicMemoryEnabled
            "启动内存"  = "$($_.MemoryStartup / 1GB)GB"
            "最小内存"  = "$($_.MemoryMinimum / 1GB)GB"
            "最大内存"  = "$($_.MemoryMaximum / 1GB)GB"
        }
        $no++
        $info
    } | Format-Table -AutoSize -Wrap | Out-Host
}
   

function Set-VMMemory() {
    param(
        [string]$VMName
    )
    if ($VMName -eq "") {
        $vm = Select-VM
        $VMName = $vm.Name
    }

    $unit = "GB"
    $uInput = Get-UserInputWithOptions  -Prompt "是否启用动态内存({{ Y }} / {{ N }}): " -Options Y, N
    $dynamicMemory = $uInput.ToUpper() -eq 'Y'

    $startUp = Get-UserInputWithOptions -Number -Prompt "输入启动内存大小({{ $unit }}): " -GT 0
    if ($dynamicMemory) {
        $min = Get-UserInputWithOptions -Number -Prompt "输入最小内存大小({{ $unit }}): " -GT 0
    }
    else {
        $min = $startUp
    }
    $max = Get-UserInputWithOptions  -Prompt "输入最大内存大小({{ $unit }}): " -GE $min
   
    Set-VM $VMName -DynamicMemory:$dynamicMemory -MemoryStartupBytes "$startUp$unit" -MemoryMinimumBytes "$min$unit" -MemoryMaximumBytes "$max$unit"
    Show-VM -VMName $VMName
}



function Select-Operation {
    $ops = @(
        [Operation]::new( "分配 Gpu 给 VM", { Set-GpuToVm }),
        [Operation]::new( "从 VM 移除 已分配 Gpu", { Remove-GpuFromVm }),
        [Operation]::new( "查看 VM 已分配 Gpu", { 
                $vm = Select-VM
                Show-VMGpu $vm.Name
            }),
        [Operation]::new( "更新 VM Gpu 驱动", { Update-VMGpuDriver }),
        [Operation]::new( "更新 GPU 资源分配占比", { Set-VMGpuResource }),
        [Operation]::new( "展示 VM 状态", { Show-VM }),
        [Operation]::new( "Get-VMGpuPartitionAdapter", { 
                $gpu = Select-VMGpu
                Get-VMGpuPartitionAdapter -VMName $gpu.VMName -AdapterId $gpu.AdapterId
            })
    )
    
    $op = Select-Item -Items $ops -ItemDescription "操作" -Properties Name

    try {
        $op.Execute()
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        exit
    }
}

function Test {
    $VMName = "xnj-1"
    $GpuInfo = [GpuInfo]::new()
    $GpuInfo.FriendlyName = "NVIDIA GeForce RTX 2080 SUPER"
    $GpuInfo.ServiceName = "nvlddmkm"
    $GpuInfo.DeviceId = "PCI\VEN_10DE&DEV_1E81&SUBSYS_1E8110DE&REV_A1\4&1FC990D7&0&0019"


    $value = Get-ItemProperty -Path "HKLM:HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\$($GpuInfo.DeviceId)"
    $value = Get-ItemProperty -Path "HKLM:HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\$($value.Driver)"

    $memorySize = $value."HardwareInformation.qwMemorySize"
    $lowMemoryMappedIoSpace = 1GB
    $highMemoryMappedIoSpace = $memorySize + 1GB

    Write-Host "设置内存映射 IO 空间：" -NoNewline
    Write-Host "Low: " -ForegroundColor Green -NoNewline
    Write-Host "$($lowMemoryMappedIoSpace / 1GB)GB" -ForegroundColor Blue -NoNewline
    Write-Host "," -NoNewline
    Write-Host "High: "-ForegroundColor Green  -NoNewline
    Write-Host "$($highMemoryMappedIoSpace/1GB)GB" -ForegroundColor Blue
    
}

Select-Operation
# Test
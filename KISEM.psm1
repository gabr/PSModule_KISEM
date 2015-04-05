# For working on Multilayer Perceptron paper for KISEM at FAU
# 20.02.2015
function Get-KISEM {
    # run LaTeX editor
    Start-Process "C:\Program Files (x86)\Texmaker\texmaker.exe"

    # change window title
    $host.ui.RawUI.WindowTitle = "KISEM"

    # run new process which will set windows positions
    Start-Process PowerShell -ArgumentList "KISEM_SetPosition"

    # navigate to working directory
    cd D:\studia\erlangen\sem5\KISEM\Praca;
    # show files
    dir;
    pwd;
}

function KISEM_SetPosition {
    # set windows position and size
    Import-Module WASP # module for managing windows

    do {
        $editor = Select-Window | where { $_.ProcessName.Equals("texmaker") } | Select -First 1
    } while ($editor -eq $null)
    $editor.Restore()
    Set-WindowPosition -X -4 -Y -4 -Window $editor
    $editor.Maximize()

    do {
        $console = Select-Window | where { $_.Title.Equals("KISEM") } | Select -First 1
    } while ($console -eq $null)
    $console.Restore()
    Set-WindowPosition -X -488 -Y 385 -Width 488 -Height 777 -Window $console

    # if there is web browser move it to
    $browsers = Select-Window | where { $_.Title.Contains("wiadomości") }

    foreach ($browser in $browsers)
    {
        $browser.Restore()
        Set-WindowPosition -X -1366 -Y 385 -Width 879 -Height 768 -Window $browser
    }

    # open project in texmaker
    $editor | Send-Keys "^o"
    Start-Sleep -Milliseconds 500
    do {
        $open = $editor | Select-ChildWindow | Select -First 1
    } while ($open -eq $null)
    $open.Activate()
    $open | Send-Keys "{F4}"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "^a"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "{DEL}"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "D:\studia\erlangen\sem5\KISEM\Praca\"
    $open | Send-Keys "{ENTER}"
    Start-Sleep -Milliseconds 500

    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"

    Start-Sleep -Milliseconds 500
    $open | Send-Keys "MLP.tex"
    Start-Sleep -Milliseconds 500
    $open | Send-Keys "{ENTER}"
    Start-Sleep -Milliseconds 500
    $editor.Activate()
    Start-Sleep -Milliseconds 200
    $editor | Send-Keys "%o"
    Start-Sleep -Milliseconds 200
    $editor | Send-Keys "{DOWN}"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "{ENTER}"

    # open Open window again to choose Sections file
    $open = $null
    Start-Sleep -Milliseconds 500
    $editor | Send-Keys "^o"
    Start-Sleep -Milliseconds 500
    do {
        $open = $editor | Select-ChildWindow | Select -First 1
    } while ($open -eq $null)
    $open.Activate()
    $open | Send-Keys "{F4}"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "^a"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "{DEL}"
    Start-Sleep -Milliseconds 200
    $open | Send-Keys "D:\studia\erlangen\sem5\KISEM\Praca\Sections\"
    $open | Send-Keys "{ENTER}"
    Start-Sleep -Milliseconds 500

    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"
    $open | Send-Keys "{F6}"

    Start-Sleep -Milliseconds 200
    $open.Activate()
}

Export-ModuleMember -Function Get-KISEM
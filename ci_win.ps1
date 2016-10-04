$ErrorActionPreference = "Stop"

# System Config
$msys = $env:MSYS
if (!$msys) {throw "MSYS environment variable not set"}

$mingw_dir = "$msys\mingw64"
$start_path = $env:Path
$env:Path += ";$mingw_dir\bin"
# we use a short stack root path to try and avoid triggering command-line-length errors on windows
$env:STACK_ROOT += "c:\sw"
$mingw_path = $env:Path

# Project Config
$dlls = @("libgcc_s_seh-1.dll", "libsodium-13.dll", "libstdc++-6.dll", "libwinpthread-1.dll")
$built_exes = @("xlsx2yaml.exe")
$project = "xlsx2yaml"
$target = "xlsx2yaml"

$packages = @("zip")

echo "----------------"
echo "Installing Pacman Package Deps"

$packages |`
  ForEach-Object {
    & $msys\usr\bin\pacman -S -q $_ --needed --noconfirm
  }

echo "----------------"
echo "Building with Stack"

& stack setup
& stack build $target --extra-include-dirs=$msys/mingw64/include --extra-lib-dirs=$msys/mingw64/lib
if (!$?) { throw "Build failed" }

echo "----------------"
echo "Packaging"

if (test-path "pkg") {
  remove-item -recurse "pkg"
}
new-item "pkg" -type directory -force | Out-Null

# Stack outputs some other gunf so look for the project path in the output here
$installdir = stack path --local-install-root | Select-String $project

echo "Install Dir"
echo "$installdir"

echo "Copying Exes"
$built_exes |`
  ForEach-Object {
    Copy-Item "$installdir\bin\$_" "pkg\$_"
  }

echo "Copying DLLs"
$dlls |`
  ForEach-Object {
    $it = "$mingw_dir\bin\$_"
    Copy-Item $it "pkg"
  }

echo "Testing Exes"

$env:Path = $start_path
$built_exes |`
  ForEach-Object {
    & "pkg\$_" --help 2>&1
    if (!$?) { throw "$_ failed to --help, maybe a dll is missing?" }
  }
$env:Path = $mingw_path

#####
echo "Creating pkg/$($project)_windows.zip"
#####

$xs = Get-ChildItem -path pkg
echo "Adding to package: $($xs -join ", ")"
Push-Location "pkg"
& $msys\usr\bin\zip -D "$($project)_windows.zip" $xs
Pop-Location

echo "----------------"
echo "Done."

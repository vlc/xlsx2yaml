$ErrorActionPreference = "Stop"

# Project Config
$dlls = @("libgcc_s_seh-1.dll", "libsodium-18.dll", "libstdc++-6.dll", "libwinpthread-1.dll")
$built_exes = @("xlsx2yaml.exe")
$project = "xlsx2yaml"
$target = "xlsx2yaml"

$packages = @("zip")

# ---------------------------------------------------

# TODO: newer versions of stack prefer '--programs'
$mingw_dir = stack path --ghc-paths
$mingw_dir = "$mingw_dir\msys2-20150512\mingw64\bin"

echo "----------------"
echo "Installing Pacman Package Deps"

# Update pacman package database. Required at least once.
& stack exec -- pacman -Syy 
$packages |`
  ForEach-Object {
    & stack exec -- pacman -S -q $_ --needed --noconfirm
  }

echo "----------------"
echo "Building"

& stack build --test
if (!$?) { throw "Build failed" }

echo "----------------"
echo "Packaging"

if (test-path "pkg") {
  remove-item -recurse "pkg"
}
new-item "pkg" -type directory -force | Out-Null

# Stack outputs some other gunf so look for the project path in the output here
$installdir = stack path --local-install-root | Select-String k

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
    $it = "$mingw_dir\$_"
    Copy-Item $it "pkg"
  }

echo "Testing Exes"

$built_exes |`
  ForEach-Object {
    & "pkg\$_" --help 2>&1
    if (!$?) { throw "$_ failed to --help, maybe a dll is missing?" }
  }

echo "Creating pkg/$($project)_windows.zip"

$xs = Get-ChildItem -path pkg
echo "Adding to package: $($xs -join ", ")"
Push-Location "pkg"
& stack exec -- zip -D "$($project)_windows.zip" $xs
Pop-Location

echo "----------------"
echo "Done."

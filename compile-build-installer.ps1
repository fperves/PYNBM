$version = get-content PYNBM.py | %{ if ($_.Split('"')[0] -match "^Version") { $_.Split('"')[1]; } }

invoke-expression "python setup.py py2exe"
$nsiscmd = 'c:\Program Files (x86)\NSIS\makensis.exe'
$nsisopt = '/DVERSION='+$version
$nsisopt2 = 'installer.nsi'
& $nsiscmd $nsisopt $nsisopt2

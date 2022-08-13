# Sources 

$urlEvkey = "https://github.com/lamquangminh/EVKey/releases/download/Release/EVKey.zip"
$urlZalo = "https://res-download-pc-te-vnno-pt-6.zadn.vn/hybrid/ZaloSetup_22.7.2.exe"
$urlOffice = "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img"

# Set up file and directory

cd C:\
mkdir AllInOne
cd .\AllInOne\
mkdir EVKey
cd .\EVKey\

$outEvkey = "C:\AllInOne\EVKey\EVKey.zip"
$outZalo = "C:\AllInOne\zalo.exe"
$outOffice = "C:\AllInOne\office.iso"

# Import BITS (https://docs.microsoft.com/en-us/windows/win32/bits/about-bits?redirectedfrom=MSDN)

Import-Module BitsTransfer

# Download and Install EVKey

echo "Downloading EVKey"
Start-BitsTransfer -Source $urlEvkey -Destination $outEvkey
echo "Download Completed"
Expand-Archive $outEvkey -DestinationPath . -Force
echo "Unzip File"
del .\EVKey.zip
.\EVKey64.exe
echo "Install EVKey Completed"
echo "-----------------------"

# Download and Install Zalo

echo "Downloading Zalo"
cd ../
Start-BitsTransfer -Source $urlZalo -Destination $outZalo
echo "Download Completed"
.\zalo.exe
echo "Install Zalo Completed"
echo "-----------------------"

# Download and Install Office 2019

Start-BitsTransfer -Source $urlOffice -Destination $outOffice
$DrivesBeforeMount = (Get-PSDrive).Name
Mount-DiskImage -ImagePath $outOffice
$DrivesAfterMount = (Get-PSDrive).Name
$DriveLetterUsed = (Compare-Object -ReferenceObject $DrivesBeforeMount -DifferenceObject $DrivesAfterMount).InputObject
Set-Location -Path "$DriveLetterUsed`:\"
.\Setup.exe
cd C:\AllInOne

# End
echo "Press any key to quit ..."
[Console]::ReadKey()

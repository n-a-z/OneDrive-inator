# Autor: Piotr Stępień
# Skrypt został napisany pod Windows 10. W wyapdku innych wersji Windowsa może nie działać poprawnie dlatego sprawdza wersję Windowsa i w wypadku wersji innej niż 10 wysyła informację zwrotną do użytkownika i ewentualnie Admina.

##### Zmienne modyfikowalne.

$OD_name = "OneDrive - Company_Name" # Nazwa folderu OneDrive. Dla kont służbowych będzie to domyślnie "OneDrive - nazwa_firmy", dla kont prywatnych "OneDrive". POLE WYMAGANE!
$OD_custom = "False" # Jeżeli ścieżka do folderu OneDrive będzie inna od domyślnej (%USERPROFILE%), wpisz ją tutaj. Zaleca się zachowanie domyślnej ścieżki i pozostawienie wartości "False".
$OD_folder = "OD_Copy" # Nazwa folderu, w którym będą przechowywane foldery użytkowników. Ma być puste ("") jeżeli mają być w folderze głównym OneDrive.
$prompt_time = 30 #minut # Jeżeli użytkownik nie ma skonfigurowanego programu OneDrive, skrypt poprosi go o zalogowanie. Określ co ile minut skrypt ma powtarzać prośbę o zalogowanie. Wartość nie może znajdować się cytacie ("przykładowa_niepoprawna_wartość")!
$OD_limit = 15 #GB # OneDrive posiada limit maksymalnego rozmiaru plików. Limit ten czasami się zmienia. NAleży tutaj wpisać aktualny limit w GB.

$move_desktop = "True" # Wybierz czy ma być przeniesiony folder Pulpit. Wpisać wartość inną niż True aby pominąć przenoszenie.
$move_documents = "True" # Wybierz czy ma być przeniesiony folder Dokumenty. Wpisać wartość inną niż True aby pominąć przenoszenie.
$move_pictures = "True" # Wybierz czy ma być przeniesiony folder Obrazy. Wpisać wartość inną niż True aby pominąć przenoszenie.
$move_video = "True" # Wybierz czy ma być przeniesiony folder Filmy. Wpisać wartość inną niż True aby pominąć przenoszenie.
$move_music = "True" # Wybierz czy ma być przeniesiony folder Muzyka. Wpisać wartość inną niż True aby pominąć przenoszenie.
$move_favorites = "True" # Wybierz czy ma być przeniesiony folder Ulubione. Wpisać wartość inną niż True aby pominąć przenoszenie.
$move_downloads = "False" # Wybierz czy ma być przeniesiony folder Pobrane. Wpisać wartość inną niż True aby pominąć przenoszenie.

$nametest = "True" # Sprawdzanie poprawności nazw (usuwa niedozwolone # oraz % z nazw plików i zastępuje je słowami "hash" oraz "prcnt". Proces zajmuje od kilku minut do godziny. Wpisać wartość inną niż True aby pominąć. 
$sizetest = "True" # Sprawdza czy są pliki większe niż $OD_limit i wypisuje je w pliku "$OD_folderk\brak_synchro.txt" oraz wysyła to samo info na adres mailowy Admina oraz na do ścieżki sieciowej jeżeli te opcje zostaną ustawione na "True". Pliki większe niż $OD_limit nie będą synchronizowane.  Wpisać wartość inną niż True aby pominąć.
$PST_Copy = "True" # Kopiuje pliki PST do oryginalnej lokacji ($Env:userprofile\Documents\Pliki Programu Outlook) aby nie kożystał z plików PST umieszczonych na OneDrivie. Wpisać wartość inną niż True aby pominąć. UWAGA, opcja jest aktywna tylko gdy $move_documents = "True".
$WinCheck = "True" # Program informuje użytkownika i opcjonalnie admina o złej wersji Windowsa. Zaleca się zachowanie wartości domyślnej "True". Wpisać wartość inną niż True aby pominąć.

$admin_error_win = "True" # Jeżeli "True" Admin będzie dostawać maile od użytkowników, którzy mają złą wersją Windowsa. Maile mogą być blokowane przez antyspam dlatego te informacje będą również zapisane na wybranej ścieżce sieciowej. Wpisać wartość inną niż True aby pominąć. Użytkownik  dostanie informacje o niezależnie od tej opcji. Działa tylko gdy $WinCheck ma wartość "True".
$admin_error_size = "True" # Jeżeli "True" Admin będzie dostawać maile od użytkowników, którzy mają za duże pliki. Maile mogą być blokowane przez antyspam dlatego te informacje będą również zapisane na wybranej ścieżce sieciowej. Wpisać wartość inną niż True aby pominąć. Użytkownik zawsze dostanie informacje o błędach.
$admin_success = "True" # Jeżeli "True" program będzie generować listę użytkowników, którzy prześli na OneDrive. Plik nosi nazwę "OneDrive_Migration.txt", znajduje się w $net_error_lok i zawiera nazwy użytkowników, nazwy komputerów, oraz datę przejścia.

# Poniższe wypełnić tylko jeżeli $admin_error /$admin_success ma wartość "True".
$net_error_lok = "\\server_address\one_drive\" # Ścieżka sieciowa, w kótrej będą trzymane informacje o błędach od użytkowników. W przypadku za dużych plików pojawi się plik tekstowy w folderze "BigFiles" z nazwą użytkownia oraz datą zawierający listę zbyt dużych plików. W wypadku złej wersji Windows, użytkownicy wraz z ich wersją Windowsa zostaną wypiani w pliku "OneDrive_old_Windows.txt". Pozostawić "False" aby pominąć.
$admin_mail = "admin@email.address" # Adres e-mail administratora, który ma dostawać ewentualne powiadomienia. Ma być puste ("") bądź fałszywe aby pominąć.
$company_smtp = "serwer_name.mail.protection.outlook.com" # serwer smtp, przez który nastąpi wysyłanie. Działają tylko serwery Direct Send.
$smtp_port = "25" # Port smtp serwera.

##### Zmienna określająca nową lokalizację folderów profilowych. NIE ZMIENIAĆ.
if ($OD_custom -ne "False")
{
$lok = "$OD_custom\$OD_name\$OD_folder"
$lokreg = "$OD_custom\$OD_name\$OD_folder"
}
else
{
$lok = "$Env:userprofile\$OD_name\$OD_folder"
$lokreg = "%USERPROFILE%\$OD_name\$OD_folder"
}

##### Zmienna określająca lokalizację w rejestrze folderów użytkowników. Wartość domyślna jest dla Windowsa 10. NIE ZMIENIAĆ.
$Regkey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

##### Poniższe komunikaty można opcjonalnie modyfikować. Wartości domyślne powinny być optymalne.

# Tekst komunikatu logowania się do programu OneDrive.
$dial1_text = "Proszę zalogować się do programu OneDrive."
# Tekst komunikatu rozpoczęcia konfiguracji programu OneDrive.
$dial2_text = "Nastąpi konfiguracja programu OneDrive.`nMoże po potrwać do kilkunatu minut.`nProszę zapisać i zamknąć wszystkie dokumenty a następnie wcisnąć przycisk 'OK'."
# Tekst komunikatu zakończenia konfiguracji programu OneDrive.
$dial3_text = "Program OneDrive został skonfigurowany.`nPliki znajdują się w lokalizacji:`n'$lok'."
# Tekst komunikatu o plikach większych niż $OD_limit.
$TooBig_text = "Na twoim komputerze znajdują się pliki zbyt duże do synchronizacji.`nIch lista znajduje się w pliku:`n'$lok\brak_synchro.txt'"
# Tekst komunikatu złej wersji Windowsa.
$WinError_text = "Aby skonfigurować program OneDrive niezbędne jest zainstalowanie Windows 10.`nProszę skontaktować się z działem IT."

##### Koniec zmiennych modyfikowalnych

##### Program sprawdza wersję Windowsa

#$WinVer = gwmi win32_operatingsystem | % caption
if ((Get-WmiObject Win32_OperatingSystem).Caption -eq "Microsoft Windows 10 Pro")
{
##### START Konfiguracja

##### Sprawdzenie czy pliki już są na OneDrive'ie
if ((Get-ItemProperty -Path $Regkey -Name Desktop).Desktop -like "*$lok*")
{
& "$Env:userprofile\AppData\Local\Microsoft\OneDrive\OneDrive.exe" /background
#$dial0 = New-Object -ComObject Wscript.Shell
#$dial0.Popup("Program OneDrive jest już skonfigurowany.`nPliki znajdują się w '$lok'",0,"$OD_name")
exit
}   
else
{

  ##### Sprawdzenie czy użytkownik jest zalogowany do OneDrive'a kontem służbowym
  $path = Test-Path "$Env:userprofile\$OD_name\"
  if ($path -ne "True") 
  {
  ##### Wyłączenie programu OneDrive
  #Start-Sleep -s 30
  Stop-Process -processname OneDrive -ErrorAction SilentlyContinue

    ##### Pętla wywołująca logowanie do programu OneDrive
    $prompt_time = $prompt_time*60
    while ($path -ne "True")
    {   
    $dial1 = New-Object -ComObject Wscript.Shell
    $dial1.Popup("$dial1_text",0,"$OD_name")
    & "$Env:userprofile\AppData\Local\Microsoft\OneDrive\OneDrive.exe"
    #exit #usuń to
    ##### Pauza aby użytkownik mógł się zalogować
    Start-Sleep -s "$prompt_time"
    
    ##### Sprawdzenie czy użytkownik jest zalogowany do OneDrive'a kontem służbowym    
    $path = Test-Path "$Env:userprofile\$OD_name\"
    }
   }

    ##### START konfiguracja
    #for ($I = 1; $I -le 100; $I++ )
    #{Write-Progress -Activity "Search in Progress" -Status "$I% Complete:" -PercentComplete $I;}     
    
    ##### Informacja dla użytkownika  
    $dial2 = New-Object -ComObject Wscript.Shell
    $dial2.Popup("$dial2_text",0,"$OD_name")
 
    ##### Wyłączenie programów
    Stop-Process -processname outlook -ErrorAction SilentlyContinue
    Stop-Process -processname lync -ErrorAction SilentlyContinue
    Stop-Process -processname winword -ErrorAction SilentlyContinue
    Stop-Process -processname excel -ErrorAction SilentlyContinue
    Stop-Process -processname powerpnt -ErrorAction SilentlyContinue
    Stop-Process -processname notepad -ErrorAction SilentlyContinue
    Stop-Process -processname PDFXCview -ErrorAction SilentlyContinue
    Stop-Process -processname FoxitReader -ErrorAction SilentlyContinue
    Stop-Process -processname AcroRd32 -ErrorAction SilentlyContinue
    Stop-Process -processname explorer
    #Start-Sleep -s 5
    #Stop-Process -processname mstsc
    
    ##### Tworzenie folderu BACKUP
    New-Item -ItemType directory -Path $lok -ErrorAction SilentlyContinue

    ##### Zamiana znaków niedozwolonych na OneDrive'ie
   if ($nametest -eq "True")
   {
    Write-Host "Sprawdzanie poprawności nazw plików.`nMoże to potrwać kilka minut."
    Get-ChildItem "$Env:userprofile\Desktop\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 9%"
    Get-ChildItem "$Env:userprofile\Documents\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 18%"
    Get-ChildItem "$Env:userprofile\Pictures\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 27%"
    Get-ChildItem "$Env:userprofile\Videos\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 36%"
    Get-ChildItem "$Env:userprofile\Music\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 45%"
    Get-ChildItem "$Env:userprofile\Favorites\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 54%"
    Get-ChildItem "$Env:userprofile\Desktop\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 63%"
    Get-ChildItem "$Env:userprofile\Documents\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 72%"
    Get-ChildItem "$Env:userprofile\Pictures\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Get-ChildItem "$Env:userprofile\Videos\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 81%"
    Get-ChildItem "$Env:userprofile\Music\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 99%"
    Get-ChildItem "$Env:userprofile\Favorites\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukończono 100%"

    ##### Wyłączenie programów
    Stop-Process -processname outlook -ErrorAction SilentlyContinue
    Stop-Process -processname lync -ErrorAction SilentlyContinue
    Stop-Process -processname winword -ErrorAction SilentlyContinue
    Stop-Process -processname excel -ErrorAction SilentlyContinue
    Stop-Process -processname powerpnt -ErrorAction SilentlyContinue
    Stop-Process -processname notepad -ErrorAction SilentlyContinue
    Stop-Process -processname PDFXCview -ErrorAction SilentlyContinue
    Stop-Process -processname FoxitReader -ErrorAction SilentlyContinue
    Stop-Process -processname AcroRd32 -ErrorAction SilentlyContinue
    Stop-Process -processname AcroRd32 -ErrorAction SilentlyContinue
    #Stop-Process -processname explorer
    #Start-Sleep -s 5
   }
   
   ##### Przenoszenie Pulpitu
   if ($move_desktop -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Desktop" -ErrorAction SilentlyContinue
    #robocopy "$Env:userprofile\Desktop" "$lok\Desktop" /e /xf *
    
    $newPathDesktop = "$lokreg\Desktop"
    #$keyDesktop2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Desktop $newPathDesktop
    set-ItemProperty -path $Regkey -name "{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}" $newPathDesktop
    #set-ItemProperty -path $keyDesktop2 -name Desktop $newPathDesktop

    #robocopy "$Env:userprofile\Desktop" "$lok\"
    #Move-Item "$Env:userprofile\Desktop\*" "$lok\Desktop\" -force
    #Get-ChildItem "$Env:userprofile\Desktop" -Recurse | Move-Item -Destination "$lok\Desktop"
    Get-ChildItem "$Env:userprofile\Desktop\*" -Force | Move-Item -Destination "$lok\Desktop\"

    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboDesk = Get-ChildItem -Path "$Env:userprofile\Desktop" | Measure-Object
    if ($RoboDesk.count -ne "0")
    {
    robocopy "$Env:userprofile\Desktop" "$lok\Desktop" /move /e
    }
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Desktop" -Recurse -Force
    }
   }

   ##### Przenoszenie Dokumentów
   if ($move_documents -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Documents" -ErrorAction SilentlyContinue
    
    $newPathDocuments = "$lokreg\Documents"
    #$keyDocuments2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Personal $newPathDocuments
    set-ItemProperty -path $Regkey -name "{F42EE2D3-909F-4907-8871-4C22FC0BF756}" $newPathDocuments
    #set-ItemProperty -path $keyDocuments2 -name Personal $newPathDocuments

    #Move-Item "$Env:userprofile\Documents\*" "$lok\Documents\" -force
    Get-ChildItem "$Env:userprofile\Documents\*" -Force | Move-Item -Destination "$lok\Documents\"
    #Get-ChildItem "$Env:userprofile\Documents" -Recurse | Move-Item -Destination "$lok\Documents"

    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboDok = Get-ChildItem -Path "$Env:userprofile\Documents" | Measure-Object
    if ($RoboDok.count -ne "0")
    {
    robocopy "$Env:userprofile\Documents" "$lok\Documents" /move /e
    }
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Documents" -Recurse -Force
    }
    # Skopiowanie plików PST do domyślnej lokalizacji
    if ($PST_Copy -eq "True")
    {
      if (Test-Path "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook")
      {
      New-Item -ItemType directory -Path "$Env:userprofile\Documents\Pliki programu Outlook" -ErrorAction SilentlyContinue
      $maxsize = $OD_limit*1000*1024*1024
      robocopy "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook" "$Env:userprofile\Documents\Pliki programu Outlook" /e /MAX:$maxsize
      Get-ChildItem "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook\*" -Force | Move-Item -Destination "$Env:userprofile\Documents\Pliki programu Outlook\" -ErrorAction SilentlyContinue
      #robocopy "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook" "$Env:userprofile\Documents\Pliki programu Outlook" /mov /e /MIN:$maxsize
      }
    }
   }

   ##### Przenoszenie Obrazów
   if ($move_pictures -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Pictures" -ErrorAction SilentlyContinue
    
    $newPathPictures = "$lokreg\Pictures"
    #$keyPictures2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name "My Pictures" $newPathPictures
    set-ItemProperty -path $Regkey -name "{0DDD015D-B06C-45D5-8C4C-F59713854639}" $newPathPictures
    #set-ItemProperty -path $keyPictures1 -name "{33E28130-4E1E-4676-835A-98395C3BC3BB}" $newPathPictures
    #set-ItemProperty -path $keyPictures2 -name "My Pictures" $newPathPictures

    #Move-Item "$Env:userprofile\Pictures\*" "$lok\Pictures\" -force
    Get-ChildItem "$Env:userprofile\Pictures\*" -Force | Move-Item -Destination "$lok\Pictures\"
    #Get-ChildItem "$Env:userprofile\Pictures" -Recurse | Move-Item -Destination "$lok\Pictures"
    
    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboPic = Get-ChildItem -Path "$Env:userprofile\Pictures" | Measure-Object
    if ($RoboPic.count -ne "0")
    {
    robocopy "$Env:userprofile\Pictures" "$lok\Pictures" /move /e
    }
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Pictures" -Recurse -Force
    }
   }
    
   ##### Przenoszenie Video
   if ($move_video -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Videos" -ErrorAction SilentlyContinue
     
    $newPathVideos = "$lokreg\Videos"
    #$keyVideos2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name "My Video" $newPathVideos
    set-ItemProperty -path $Regkey -name "{35286A68-3C57-41A1-BBB1-0EAE73D76C95}" $newPathVideos
    #set-ItemProperty -path $keyVideos2 -name "My Video" $newPathVideos

    #Move-Item "$Env:userprofile\Videos\*" "$lok\Videos\" -force
    Get-ChildItem "$Env:userprofile\Videos\*" -Force | Move-Item -Destination "$lok\Videos\"
    #Get-ChildItem "$Env:userprofile\Videos" -Recurse | Move-Item -Destination "$lok\Videos"

    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboVid = Get-ChildItem -Path "$Env:userprofile\Videos" | Measure-Object
    if ($RoboVid.count -ne "0")
    {
    robocopy "$Env:userprofile\Videos" "$lok\Videos" /move /e
    }    
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Videos" -Recurse -Force
    }
   }
        
   ##### Przenoszenie Muzyki
   if ($move_music -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Music" -ErrorAction SilentlyContinue
    
    $newPathMusic = "$lokreg\Music"
    #$keyMusic2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name "My Music" $newPathMusic
    set-ItemProperty -path $Regkey -name "{A0C69A99-21C8-4671-8703-7934162FCF1D}" $newPathMusic  
    #set-ItemProperty -path $keyMusic2 -name "My Music" $newPathMusic
    
    #Move-Item "$Env:userprofile\Music\*" "$lok\Music\" -force
    Get-ChildItem "$Env:userprofile\Music\*" -Force | Move-Item -Destination "$lok\Music\"
    #Get-ChildItem "$Env:userprofile\Music" -Recurse | Move-Item -Destination "$lok\Music"

    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboMuz = Get-ChildItem -Path "$Env:userprofile\Music" | Measure-Object
    if ($RoboMuz.count -ne "0")
    {
    robocopy "$Env:userprofile\Music" "$lok\Music" /move /e
    }  
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Music" -Recurse -Force
    }
   }
    
   ##### Przenoszenie Ulubionych
   if ($move_favorites -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Favorites" -ErrorAction SilentlyContinue
    
    $newPathFavorites = "$lokreg\Favorites"
    #$keyFavorites2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Favorites $newPathFavorites  
    #set-ItemProperty -path $keyFavorites2 -name Favorites $newPathFavorites

    #Move-Item "$Env:userprofile\Favorites\*" "$lok\Favorites\" -force
    Get-ChildItem "$Env:userprofile\Favorites\*" -Force | Move-Item -Destination "$lok\Favorites\"
    #Get-ChildItem "$Env:userprofile\Favorites" -Recurse | Move-Item -Destination "$lok\Favorites"

    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboFav = Get-ChildItem -Path "$Env:userprofile\Favorites" | Measure-Object
    if ($RoboFav.count -ne "0")
    {
    robocopy "$Env:userprofile\Favorites" "$lok\Favorites" /move /e
    } 
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Favorites" -Recurse -Force
    }
   }

   ##### Przenoszenie Pobranych
   if ($move_downloads -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Downloads" -ErrorAction SilentlyContinue
    
    $newPathDownloads = "$lokreg\Downloads"
    #$keyDownloads2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Downloads $newPathDownloads
    set-ItemProperty -path $Regkey -name "{374DE290-123F-4565-9164-39C4925E467B}" $newPathDownloads
    #set-ItemProperty -path $keyDownloads2 -name Downloads $newPathDownloads

    Get-ChildItem "$Env:userprofile\Downloads\*" -Force | Move-Item -Destination "$lok\Downloads\"

    ##### Sprawdzenie czy wszystko zostało przeniesione
    $RoboDow = Get-ChildItem -Path "$Env:userprofile\Downloads" | Measure-Object
    if ($RoboDow.count -ne "0")
    {
    robocopy "$Env:userprofile\Downloads" "$lok\Downloads" /move /e
    } 
    ##### Usunięcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Downloads" -Recurse -Force
    }
   }

   ##### Sprawdzenie plików większych niż $OD_limit
   if ($sizetest -eq "True")
   {
    $OD_limit = "$OD_limit"+"GB"
    $BigFiles = Get-ChildItem "$Env:userprofile\" -Recurse | Where-Object {$_.Length -gt "$OD_limit"}

   if ($BigFiles) 
   {    
    $date = (Get-Date).ToShortDateString()
    $BigFiles >> "$lok\brak_synchro.txt"
    #$BigFiles | Set-Content "$lok\brak_synchro.txt" #same nazwy, bez ścieżek
    
    ##### Wysłanie maila z listą plików i nazwą użytkownika do admina oraz zapisanie tej informacji na ścieżce sieciowej.
    if ($admin_error_size -eq "True")
    {
    Send-MailMessage -To "$admin_mail" -From "$admin_mail" -Subject "OneDrive - $env:UserName - za duze pliki" -Body "$env:UserName`n$BigFiles" -SmtpServer "$company_smtp" -Port "$smtp_port" -Attachments "$lok\brak_synchro.txt" -ErrorAction SilentlyContinue
      if (Test-Path "$net_error_lok")
      {
      $date = (Get-Date).ToShortDateString()
      New-Item -ItemType directory -Path "$net_error_lok\BigFiles" -ErrorAction SilentlyContinue
      $BigFiles >> "$net_error_lok\BigFiles\$date;$env:UserName.txt"
      }
    }
    $TooBig = New-Object -ComObject Wscript.Shell
    $TooBig.Popup("$TooBig_text",0,"$OD_name")
   }
   }
   
    ##### Zapisanie w $net_error_lok\OneDrive_Migration.txt daty, nazwy użytkownika oraz komputera, któremu udało się przenieść pliki na OneDrive'a.
    if ($admin_success -eq "True")
    {
     $date = (Get-Date).ToShortDateString()
     New-Item -ItemType directory -Path "$net_error_lok" -ErrorAction SilentlyContinue
     "$date;$env:computername;$env:UserName" >> "$net_error_lok\OneDrive_Migration.txt"
    }
    ##### Restart Explorera i komunikat końca
    Stop-Process -processname explorer
    #Start-Process explorer
    $dial3 = New-Object -ComObject Wscript.Shell
    $dial3.Popup("$dial3_text",0,"$OD_name")
exit
}
##### END konfiguracja 
}

elseif ($WinCheck = "True")
##### Program informuje o złej wersji Windowsa
{
 if ($admin_error_win -eq "True")
 {
  $winpath = Test-Path "$Env:userprofile\WinVer.dat"
  if ($winpath -ne "True")
  {
    if (Test-Path "$net_error_lok")
    {
    $date = (Get-Date).ToShortDateString()
    New-Item -ItemType directory -Path "$net_error_lok" -ErrorAction SilentlyContinue
    "$date;$env:computername;$env:UserName;$WinVer" >> "$net_error_lok\OneDrive_old_Windows.txt"
    }
  Send-MailMessage -To "$admin_mail" -From "$admin_mail" -Subject "OneDrive - $env:UserName - brak Windowsa 10" -Body "$env:UserName`n$WinVer" -SmtpServer "$company_smtp" -Port "$smtp_port" -ErrorAction SilentlyContinue
  "" >> "$Env:userprofile\WinVer.dat"
  }
 }
 $WinError = New-Object -ComObject Wscript.Shell
 $WinError.Popup("$WinError_text",0,"$OD_name")
 exit
}

else {exit}

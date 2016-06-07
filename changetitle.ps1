# Set initial values before running batch
$files = Get-ChildItem -Filter *.mp3 -Path e:\Downloads\RedRising
$bookname = 'RedRising'
$author = 'Pierce Brown'
$booktitle = 'Red Rising'
$i = 1;
$firstPR = $true

# Load TagLib for editing metadata
[Reflection.Assembly]::LoadFrom('E:\Powershell\taglib\Libraries\taglib-sharp.dll') | Out-Null

foreach($file in $files) {
   # Gather info by splitting strings
   $filename      = $file.FullName
   $fileshortname = $file.Name.Split(".")[0]
   $bookid        = $fileshortname.Split("_")[0]
   $filenumber    = $fileshortname.Split("_")[1]
   Write-Host "Modifying $fileshortname to ${bookname}_${filenumber}"

   # Load the file using TagLib
   $tagfile = [TagLib.File]::Create($file.FullName)
   
   if($filenumber -like "*PRE*") {
      $tagfile.Tag.Title = 'The Preface'
      $booktype = '_'
   }
   else {
      $booktype = $fileshortname.Split("_")[2]
   }
   
   $tagfile.Tag.Track = $i
   $tagfile.Tag.AlbumArtists = $author
   $tagfile.Tag.Artists = $author
   $tagfile.Tag.Album = $booktitle
   $tagfile.Tag.Genres = 'Audiobook'

   if($booktype -like "*IN*") {
      $tagfile.Tag.Title = 'Intro'
   }
   elseif($booktype -like "*C*") {
      $chapter = $booktype.Split("C")[1]
      $tagfile.Tag.Title = "Chapter ${chapter}"
   }
   elseif($booktype -like "*PR*") {
      $tagfile.Tag.Title = 'Prologue'
      if($firstPR) {
         $firstPR = $false
      } else {
         $i++
      }
   }
   elseif($booktype -like "*AN*") {
      $tagfile.Tag.Title = "Author's Notes"
   }
   elseif($booktype -like "*AP*") {
      $tagfile.Tag.Title = 'Appendix'
   }
   elseif($booktype -like "*EP*") {
      $tagfile.Tag.Title = 'EP'
   }
   elseif($booktype -like "*FW*") {
      $tagfile.Tag.Title = 'FW'
   }
   $tagfile.Save()

   Start-Sleep -s 1
   # Finally determine the new name and rename the file
   if($booktype -like "*IN*") {
      $newname = "${bookname}_${filenumber}_Intro.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($booktype -like "*C*") {
      $chapter = $booktype.Split("C")[1]
      $newname = "${bookname}_${filenumber}_Ch${chapter}.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($booktype -like "*PR*") {
      $newname = "${bookname}_${filenumber}_Prologue.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($filenumber -like "*PRE*") {
      $newname = "${bookname}_ThePreface.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($booktype -like "*AN*") {
      $newname = "${bookname}_${filenumber}_AuthorsNotes.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($booktype -like "*AP*") {
      $newname = "${bookname}_${filenumber}_Appendix.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($booktype -like "*EP*") {
      $newname = "${bookname}_${filenumber}_EP.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   elseif($booktype -like "*FW*") {
      $newname = "${bookname}_${filenumber}_FW.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   else {
      $newname = "${bookname}_${filenumber}_Other.mp3"
      Rename-Item -path $filename -NewName $newname
   }
   $i++
}
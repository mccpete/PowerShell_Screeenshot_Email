$OL = New-Object -ComObject outlook.application

Start-Sleep 5

<#
olAppointmentItem
olContactItem
olDistributionListItem
olJournalItem
olMailItem
olNoteItem
olPostItem
olTaskItem
#>

#Create Item
$mItem = $OL.CreateItem("olMailItem")

$mItem.To = "*******"

$mItem.Subject = "AV"
#need to have input for what is being confirmed
$email_body = read-host "Enter confirmation group with corresponding company"

write-host $email_body

$mItem.Body = $email_body

#need to have image be attached maybe by get-clipboard

#Bitmap might not work switch to jpg???
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
$Width = $Screen.Width
$Height = $Screen.Height
$Left = $Screen.left
$Top = $Screen.Top

$bitmap = New-Object System.Drawing.Bitmap $width, $Height

$graphic = [System.Drawing.Graphics]::FromImage($bitmap)

$graphic.CopyFromScreen($Left,$Top,0,0,$bitmap.Size)

$bitmap.Save("C:\Users\*****\OneDrive - Johnson Controls\Documents\AV April_22\imgproof\testimg.bmp")
#figure out random loop with file names
$IMG_path = ("C:\Users\*****\OneDrive - Johnson Controls\Documents\AV April_22\imgproof\testimg.bmp")
$mitem.Attachments.Add($IMG_path) 

$mItem.Send()



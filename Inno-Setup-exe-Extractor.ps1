Add-Type -AssemblyName System.Windows.Forms

#Use Windows Forms to open a file select dialog

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Windows Packages (*.msi)|*.msi'
}

$Out = $FileBrowser.ShowDialog() #Display the dialog

#Select output directory

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    Description = 'Output'
}

$Out = $FolderBrowser.ShowDialog() #Display the dialog

$FolderBrowser.SelectedPath #Variable stuff

innoextract --extract $FileBrowser.FileName $($FolderBrowser.SelectedPath)  # Make sure innoextract is added to Path!

#A helpful message

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("This is a repurposed MSI extractor PS1 I made, its not garrenteed to work. If you have any issues make a issue on Github.", 0, "Thank you for using Simple exe Extractor", 0)
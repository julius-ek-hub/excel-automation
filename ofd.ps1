[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | out-null 
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
$OpenFileDialog.initialDirectory = "C:\Users\juliu\Downloads" 
$OpenFileDialog.filter = "Excel workbooks (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv" 
$OpenFileDialog.ShowDialog() 
$OpenFileDialog.filename 

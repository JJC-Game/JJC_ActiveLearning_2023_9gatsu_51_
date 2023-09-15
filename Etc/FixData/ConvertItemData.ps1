$scriptPath = Get-Location
$excelFileName = "FixData.xlsm"
$excelFunctionName = "OutputItemFixData"
$excelPath = Join-Path $scriptPath $excelFileName
if(Test-Path $excelPath){
    $writeString = $excelFileName + "�̊֐�" + $excelFunctionName + "���Ăяo���Ă��܂�."
    Write-Output $writeString

    # Excel�I�u�W�F�N�g���擾
    $excel = New-Object -ComObject Excel.Application
    try
    {
        # Excel�t�@�C����OPEN
        $book = $excel.Workbooks.Open($excelPath)
        # �v���V�[�W�������s
        $excel.Run($excelFunctionName)
        # Excel�t�@�C����CLOSE
        $book.Close()
    }
    catch
    {
        $ws = New-Object -ComObject Wscript.Shell
        $ws.popup("�G���[ : " + $PSItem)
    }
    finally
    {
        # Excel���I��
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null

        $writeString = $excelFileName + "�̊֐�" + $excelFunctionName + "�𐳏�I��."
        Write-Output $writeString
    }

    #���ʕ���utf-8�ɕϊ�����.
    $sourceFileName = "ItemFixData.csv"
    $sourcePath = Join-Path $scriptPath $sourceFileName
    $allText = Get-Content $sourcePath -Encoding default
    Write-Output $allText | Out-File $sourcePath -Encoding UTF8
    $writeString = $sourceFileName + "��Shift_JIS����utf-8(BOM�t��)�ɕϊ����܂���."
    Write-Output $writeString

    #���ʕ���Asset�ȉ��Ɉړ�.
    $destPath = Join-Path $scriptPath "..\..\Assets\Resources\FixData"
    if(Test-Path $destPath){
        Move-Item -Path $sourcePath -Destination $destPath -Force
        $writeString = $sourceFileName + "��" + $destPath + "�ɔz�u���܂���."
        Write-Output $writeString
    }else{
        $writeString = $destPath + "������܂���. ERROR!!!!!"
        Write-Error $writeString
        Read-Host "Enter�L�[�ŏI��"
    }
    

}else{
    $writeString = $excelPath + "������܂���. ERROR!!!!!"
    Write-Error $writeString
    Read-Host "Enter�L�[�ŏI��"
}
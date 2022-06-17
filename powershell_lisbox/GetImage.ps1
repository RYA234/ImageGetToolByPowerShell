############################
#
#RYA234
#
#スクショ効率化スクリプト
############################

# Create the form
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

#グローバル変数
$strAry = New-Object System.Collections.Generic.List[string]
$strCurrentPath=Split-Path $MyInvocation.MyCommand.Path
Split-Path -Parent $scriptPath
$strSettingFileName=$strCurrentPath + "\data.txt"
$constImageName=0
$constImageDirectory=1
$constDescriptionNum=2

$strSaveDirectory = Get-Content $strSettingFileName | Select-Object -First 1
$cSampleForm = New-Object System.Windows.Forms.Form
$cListView = New-Object System.Windows.Forms.ListView

#lリストボックス設定
function ListViewSetting()
{
    $cListView.View = 'Details'
    $cListView.Width = 500
    $cListView.Height = 300
    $cListView.Columns[0].AutoResize($cSampleForm.Size())
    $cListView.Columns.Add('画像名',100)
    $cListView.Columns.Add('フォルダ',100)
    $cListView.Columns.Add('説明',200)
}

#テキストファイルの情報をリストボックスに渡す
function ListViewLoad()
{
    foreach ($l in Get-Content $strSettingFileName | Select-Object -Skip 2 )
    {
        $strListInput = $l.Split(" ")
        $strAry.Add($strListInput)
        $ListNum = $strListInput[0]
        $item = New-Object System.Windows.Forms.ListViewItem($ListNum)
        $item.SubItems.Add($strListInput[1])
        $item.SubItems.Add($strListInput[2])
        $cListView.Items.AddRange(($item))
    }
}

#フォーム設定
function FormSetting()
{
    $cSampleForm.Width = 550
    $cSampleForm.Height = 500
    $cSampleForm.Controls.Add($cListView)
    # OKボタンの設定
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(50,400)
    $OKButton.Size = New-Object System.Drawing.Size(200,30)
    $OKButton.Text = "選択した画像を取得"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cSampleForm.AcceptButton = $OKButton
    $cSampleForm.Controls.Add($OKButton)
}

#↓↓↓↓↓↓↓以下エントリーポイント↓↓↓↓↓↓↓↓#
Set-Clipboard # クリップボードの中身をクリア
ListViewSetting
ListViewLoad
FormSetting
while(1)#終了するまで実行
{
  $cImage = Get-Clipboard -Format Image # クリップボードから画像取得
     if($cImage) #クリップボードに画像が存在したら実行
     {
        $result= $cSampleForm.ShowDialog()
         # 選択後、OKボタンが押された場合、選択項目を表示
         if ($result -eq "OK")
         {
             $cClipboardImage = [Windows.Forms.Clipboard]::GetImage()
             if($cListView.SelectedItems[0].SubItems[$constImageName].Text.Contains("*")) #ワイルドカード処理
             {
                $strReplace = [Microsoft.VisualBasic.Interaction]::InputBox("使えない記号：$ () * + . [] ? \ / ^{} |", "文字を入力してください")  
                $cListView.SelectedItems[0].SubItems[$constImageName].Text = ($cListView.SelectedItems[0].SubItems[$constImageName].Text -replace "\*",$strReplace )   
             }


             $strListBoxSelectedPath=$strCurrentPath + "\output\" + $strSaveDirectory + "\" +$cListView.SelectedItems[0].SubItems[$constImageDirectory].Text
            If(!(test-path $strListBoxSelectedPath ))
             {
                   New-Item -ItemType Directory -Force -Path $strListBoxSelectedPath   #フォルダが存在しない場合新規作成
             }

             $strSaveFileName= $strListBoxSelectedPath+"\"+$cListView.SelectedItems[0].SubItems[$constImageName].Text+".jpg"
             if(test-path $strSaveFileName)
             {
                
                [System.Windows.Forms.MessageBox]::Show("ファイルが既に存在しています。")
             }
             else
             {
                $cClipboardImage.Save($strSaveFileName)  #保存する
             }
             
             Set-Clipboard # クリップボードの中身をクリア
            # $MessageBody1 ="保存先フォルダ：output\" + $cListView.SelectedItems[0].SubItems[$constImageDirectoryNum].Text 
             #$MessageBody2 ="保存した画像名："+$cListView.SelectedItems[0].SubItems[1].Text+".jpg"
            # [System.Windows.Forms.MessageBox]::Show($MessageBody1 +"`n"+ $MessageBody2)
            Invoke-Item ($strCurrentPath + "\output\" + $strSaveDirectory + "\" +$cListView.SelectedItems[0].SubItems[$constImageDirectory].Text)
         }
         else
         {
            exit
         }
     }
}



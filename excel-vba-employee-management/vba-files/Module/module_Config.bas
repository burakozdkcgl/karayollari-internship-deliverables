Attribute VB_Name = "module_Config"
' =========================================================================
' MODULE: module_Config
' Description: Stores layout constants, DB mapping, and centralized UI texts.
' =========================================================================

' Layout Coordinates & Dimensions
Public Const PANEL_LEFT As Single = 60
Public Const PANEL_TOP As Single = 30
Public Const PANEL_WIDTH As Single = 700
Public Const PANEL_HEIGHT As Single = 350
Public Const ENGINEER_WIDTH As Double = 92
Public Const ENGINEER_HEIGHT As Double = 80

' Color Constants
Public Const COLOR_GREEN As Long = 3289134  ' RGB(46, 204, 113)
Public Const COLOR_RED As Long = 3943655    ' RGB(231, 76, 60)
Public Const COLOR_WHITE As Long = 16777215 ' RGB(255, 255, 255)
Public Const COLOR_GRAY_BG As Long = &HF0F0F0

' Database Sheet & Column Mapping
Public Const MESAI_DB_SHEET_NAME As String = "mesai_db"
Public Const MESAI_DB_PERSONNEL_COLUMN As String = "A"
Public Const MESAI_DB_PERSONAL_NAME_COLUMN As String = "B"
Public Const MESAI_DB_START_DATE_COLUMN As String = "G"
Public Const MESAI_DB_DESERVED_OFF_DAY_COLUMN As String = "K"
Public Const MESAI_DB_CURRENT_QUOTA_COLUMN As String = "N"
Public Const MESAI_DB_USED_OFF_DAY_COLUMN As String = "L"
Public Const MESAI_DB_OFF_DAY_DESERVE_DATE_COLUMN As String = "J"

' --- UI TEXTS & MESSAGES ---
Public Const UI_TITLE_ASSETS As String = "zimmet bilgileri"
Public Const UI_BTN_BACK As String = "Geri"
Public Const UI_BTN_ASSET_CAPTION As String = "Zimmet" & vbCrLf & "Bilgileri"
Public Const UI_BTN_PERSONNEL_CAPTION As String = "Personel" & vbCrLf & "Bilgileri"
Public Const UI_LBL_PERSONNEL_ID As String = "Personel ID:"
Public Const UI_BTN_VERIFY As String = "Onayla & Devam"

' Verified Panel Texts
Public Const UI_VERIFIED_HELLO As String = "Merhaba, "
Public Const UI_VERIFIED_PREFIX As String = "Yıl bazlı toplam "
Public Const UI_VERIFIED_SUFFIX As String = " saat fazla mesai yapmış bulunmaktasınız."
Public Const UI_VERIFIED_WARN_LIMIT As String = "Yasal sınıra yaklaştınız."
Public Const UI_VERIFIED_THANKS As String = "Teşekkürler."
Public Const UI_TEXT_NOT_SPECIFIED As String = "[Belirtilmemiş]"

' Message Box & Error Strings
Public Const MSG_ERR_DB_NOT_FOUND_TITLE As String = "Error"
Public Const MSG_ERR_DB_NOT_FOUND_BODY As String = "Database sheet 'mesai_db' not found!"
Public Const MSG_ERR_VERIFY_FAIL_TITLE As String = "Verification Failed"
Public Const MSG_ERR_VERIFY_FAIL_BODY As String = "Invalid Personnel ID! Please make sure the ID exists in database."
Public Const MSG_ERR_SHAPE_NOT_FOUND_TITLE As String = "Error"
Public Const MSG_ERR_SHAPE_NOT_FOUND_BODY As String = "Image named '%s' was not found on the active sheet!"

' Build personnel verified info text template
Public Function BuildVerifiedInfoText(ByVal startDate As String, ByVal deserveDate As String, _
                                      ByVal deservedDays As Long, ByVal usedDays As Long, _
                                      ByVal currentQuota As Double) As String
    BuildVerifiedInfoText = startDate & " tarihinde işe başladınız." & vbCrLf & vbCrLf & _
                            deserveDate & " tarihinde hakediş gün sayısı (" & deservedDays & ") güncellenecektir." & vbCrLf & vbCrLf & _
                            "Güncel yıllık izin kotanız yıl içerisinde " & usedDays & " izin kullandığınız için " & currentQuota & " kalmıştır."
End Function
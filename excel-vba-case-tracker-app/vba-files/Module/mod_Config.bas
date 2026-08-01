Attribute VB_Name = "mod_Config"
Option Explicit

'*****************************************************************************
' Module: mod_Config
' Purpose: This module contains global configuration settings, constants, and variables for the application.
'*****************************************************************************

' =========================================================================
' 0. GLOBAL STATE VARIABLES
' =========================================================================
Public SelectedPreviousMonth As Integer
Public SelectedPreviousYear  As Integer

' =========================================================================
' 1. DIMENSIONS & LAYOUT CONFIGURATION
' =========================================================================
Public Const PANEL_LEFT   As Single = 60
Public Const PANEL_TOP    As Single = 30
Public Const PANEL_WIDTH  As Single = 700
Public Const PANEL_HEIGHT As Single = 350

' =========================================================================
' 2. COLOR PALETTE
' =========================================================================
Public Const COLOR_GREEN   As Long = 3289134    ' RGB(46, 204, 113)
Public Const COLOR_RED     As Long = 3943655    ' RGB(231, 76, 60)
Public Const COLOR_WHITE   As Long = 16777215   ' RGB(255, 255, 255)
Public Const COLOR_GRAY_BG As Long = &HF0F0F0   ' RGB(240, 240, 240)

' =========================================================================
' 3. WORKSHEET NAMES & LOCATIONS
' =========================================================================
Public Const MAIN_SHEET_NAME           As String = "Yenile"
Public Const PREVIOUS_MONTH_SHEET_NAME As String = "GECENAY"

' Database Sheet Names
Public Const DATABASE_SHEET_NAME_ARIZA As String = "ARIZA"
Public Const DATABASE_SHEET_NAME_BAKIM As String = "BAKIM"

' Incident Case Sheets
Public Const OPEN_CASES_SHEET_NAME     As String = "1 Gün Önce Açılan Vakalar"
Public Const CLOSED_CASES_SHEET_NAME   As String = "1 Gün Önce Kapanan Vakalar"

' Maintenance Sheets
Public Const OPEN_MAINT_SHEET_NAME     As String = "1 Gün Önce Açılan P. Bakımlar"
Public Const CLOSED_MAINT_SHEET_NAME   As String = "1 Gün Önce Kapanan P. Bakımlar"

' Previous Month Control Sheets
Public Const SHEET_NAME_ALL_CASES      As String = "Tüm Vakalar"
Public Const SHEET_NAME_ALL_MAINT      As String = "Tüm Periyodik Bakımlar"

Public Const OPEN_SHEET_FIRST_CASE_CELL As String = "A4"

' =========================================================================
' 4. DATABASE MAPPING & COLUMN SETTINGS
' =========================================================================
Public Const DB_OSM_DATES_START_COL As String = "A"
Public Const DB_OSM_DATES_END_COL   As String = "AL"

Public Const DB_OPEN_CELL_COL     As String = "D"
Public Const DB_OPEN_CELL_CONTEXT As String = "AÇIK VAKA SAYISI"

Public Const DB_CLOSED_NAME_COL              As String = "B"
Public Const CLOSED_CASE_SHEET_NAME_COL      As String = "K"
Public Const CLOSED_CASE_SHEET_FIRST_CELL    As String = "K4"

Public Const CLOSED_MAINT_SHEET_NAME_COL     As String = "L"
Public Const CLOSED_MAINT_SHEET_FIRST_CELL   As String = "L4"

' =========================================================================
' 5. UI LABELS & TITLES
' =========================================================================
Public Const LBL_OSM_TITLE                As String = "ARIZA & BAKIM GÜNLÜK RAPOR"
Public Const LBL_OPEN_TITLE               As String = "AÇIK KAYITLAR LİSTESİ"
Public Const LBL_CLOSED_TITLE             As String = "KAPALI KAYITLAR LİSTESİ"
Public Const LBL_PREVIOUS_MONTH_TITLE     As String = "TÜM VAKALAR RAPORU KONTROLÜ"
Public Const LBL_SELECT_PREV_MONTH_TITLE  As String = "AY & YIL SEÇİMİ"
Public Const LBL_SELECT_MONTH_PROMPT      As String = "Ay Seçiniz:"
Public Const LBL_SELECT_YEAR_PROMPT       As String = "Yıl Giriniz:"

' =========================================================================
' 6. BUTTON CAPTIONS
' =========================================================================
Public Const BTN_CAPTION_OSM_ANALYSIS           As String = "OSM Analizi"
Public Const BTN_CAPTION_BACK                   As String = "GERI DON"
Public Const BTN_CAPTION_CHECK_SHEETS           As String = "Sayfaları Kontrol Et"
Public Const BTN_CAPTION_NEXT                   As String = "ILERLE"
Public Const BTN_CAPTION_UPDATE                 As String = "GUNCELLE"
Public Const BTN_CAPTION_UPDATED                As String = "GUNCELLENDI"
Public Const BTN_CAPTION_PREVIOUS_MONTH_UPDATE   As String = "Tüm Vakalar Analizi"

' =========================================================================
' 7. PREVIOUS MONTH CONFIGURATIONS
' =========================================================================
Public Const PREV_MONTH_SEARCH_ROW_CONTEXT    As String = "KAPALI VAKA SAYISI"
Public Const PREV_MONTH_OPEN_SEARCH_CONTEXT   As String = "AÇIK VAKA SAYISI"
Public Const PREV_MONTH_SEARCH_COL            As String = "C"
Public Const PREV_MONTH_NAME_COL              As String = "B"
Public Const PREV_MONTH_HEADER_ROW            As Long = 1
Public Const PREV_MONTH_START_COL             As String = "A"
Public Const PREV_MONTH_END_COL               As String = "AL"

Public Const ALL_CASES_ASSIGNED_COL        As String = "K"
Public Const ALL_CASES_CLOSE_DATE_COL      As String = "W"
Public Const ALL_CASES_STATUS_COL          As String = "AF"

Public Const ALL_MAINT_ASSIGNED_COL        As String = "L"
Public Const ALL_MAINT_CLOSE_DATE_COL      As String = "W"
Public Const ALL_MAINT_STATUS_COL          As String = "AE"

Public Const VALID_CLOSED_STATUSES         As String = "Çözüldü,Kapandı,Kapalı,Cozuldu,KAPANDI,KAPALI"
Public Const VALID_OPEN_STATUSES           As String = "Açık,Atandı,Beklemede,İşlemde,Acik,Atandi,Islemde"

Public Const LBL_PREV_MONTH_CLOSED_TITLE   As String = "TÜM KAPALI VAKALAR ÖZETİ"
Public Const LBL_PREV_MONTH_OPEN_TITLE     As String = "TÜM AÇIK VAKALAR ÖZETİ"

Public Const LST_HDR_TYPE                  As String = "TÜR"
Public Const LST_HDR_CLOSED_COUNT          As String = "KAPALI VAKA SAYISI"
Public Const LST_HDR_OPEN_COUNT            As String = "AÇIK VAKA SAYISI"

Public Const LST_HDR_PREV_CLOSED_COUNT     As String = "KAPALI VAKA SAYISI"
Public Const LST_HDR_PREV_OPEN_COUNT       As String = "AÇIK VAKA SAYISI"

' =========================================================================
' 8. SERVICE & REPORT CATEGORY CONSTANTS
' =========================================================================
Public Const RPT_TYPE_CASE_REPORT As String = "VAKA RAPORU"
Public Const RPT_TYPE_MAINT_REPORT As String = "BAKIM RAPORU"

Public Const RPT_TYPE_OPEN_CASE    As String = "OPEN_CASE"
Public Const RPT_TYPE_CLOSED_CASE  As String = "CLOSED_CASE"
Public Const RPT_TYPE_OPEN_MAINT   As String = "OPEN_MAINT"
Public Const RPT_TYPE_CLOSED_MAINT As String = "CLOSED_MAINT"

Public Const CATEGORY_ARIZA        As String = "ARIZA"
Public Const CATEGORY_BAKIM        As String = "BAKIM"

' =========================================================================
' 9. LISTBOX HEADERS & STATUS VALUES
' =========================================================================
Public Const STATUS_NOT_FOUND As String = "NOT_FOUND"
Public Const STATUS_NO_DB_SHEET As String = "NO_DB_SHEET"
Public Const STATUS_ERROR     As String = "HATA"
Public Const STATUS_OK        As String = "OK"

Public Const CASE_TYPE_OPEN_STR        As String = "ACILAN ARIZA"
Public Const CASE_TYPE_CLOSED_STR      As String = "KAPANAN ARIZA"
Public Const CASE_TYPE_MAINT_OPEN_STR  As String = "ACILAN BAKIM"
Public Const CASE_TYPE_MAINT_CLOSED_STR As String = "KAPANAN BAKIM"

Public Const LST_HDR_STATUS        As String = "DURUM"
Public Const LST_HDR_REPORT_TYPE   As String = "RAPOR TURU"
Public Const LST_HDR_REPORT_DATE   As String = "RAPOR TARIHI"
Public Const LST_HDR_AFFECTED_DATE As String = "ETKILENEN TARIH"
Public Const LST_HDR_SHEET_NAME    As String = "SAYFA ISMI"
Public Const LST_HDR_CASE_COUNT    As String = "KAYIT SAYISI"
Public Const LST_HDR_AFFECTED_CELL As String = "ETKILENEN HUCRE"
Public Const LST_HDR_PERSON_INFO   As String = "KISI BILGISI"
Public Const LST_HDR_CASE          As String = "ADET"
Public Const LST_HDR_CELL          As String = "HUCRE"

Public Const LST_VAL_MISSING_SHEET As String = "SAYFA EKSIK!"

' =========================================================================
' 10. USER MESSAGES & SYSTEM ERRORS
' =========================================================================
Public Const MSG_TITLE_INFO      As String = "Bilgilendirme"
Public Const MSG_TITLE_SUCCESS   As String = "İşlem Başarılı"
Public Const MSG_TITLE_COMPLETED As String = "İşlem Tamamlandı"
Public Const MSG_TITLE_ERROR     As String = "Hata"

Public Const MSG_NO_SHEETS_TO_PROCESS            As String = "İşlenecek herhangi bir rapor sayfası bulunamadı."
Public Const MSG_OPEN_SAVE_SUCCESS               As String = "Açık kayıt sayıları veritabanına başarıyla kaydedildi."
Public Const MSG_CLOSED_SAVE_SUCCESS             As String = "Kapalı kayıt sayıları veritabanına başarıyla kaydedildi."
Public Const MSG_NO_MORE_CLOSED_CASES            As String = "Aktarılacak başka kayıt bulunamadı. İşlem süreci tamamlandı."
Public Const MSG_ALL_OPERATIONS_COMPLETED        As String = "Tüm güncelleme işlemleri başarıyla tamamlandı!"
Public Const MSG_DB_NOT_FOUND                    As String = "Hedef veritabanı sayfası çalışma kitabında bulunamadı!"
Public Const MSG_PREVIOUS_MONTH_SHEETS_NOT_FOUND As String = "Aranan sayfalar bulunamadı: Tüm Vakalar/Tüm Periyodik Bakımlar"
Public Const MSG_INVALID_YEAR_INPUT               As String = "Lütfen geçerli 4 haneli bir yıl giriniz! (Örn: 2026)"
Public Const MSG_PREV_MONTH_SELECTED_INFO         As String = "Seçilen Geçen Ay Raporu: "
Public Const MSG_PREV_MONTH_PROCESS_START         As String = "Tüm vakalar güncelleme işlemi başlatılıyor..."

Public Const MSG_PREV_MONTH_SAVE_SUCCESS         As String = "Tüm kapalı vaka sayıları güncellendi."
Public Const MSG_PREV_MONTH_OPEN_SAVE_SUCCESS    As String = "Tüm açık vaka sayıları güncellendi."

Public Const MSG_ERR_TARGET_DB_NOT_FOUND         As String = " veritabanı sayfası bulunamadı!"
Public Const MSG_ERR_PREV_MONTH_SHEET_NOT_FOUND  As String = " sayfası bulunamadı!"

' =========================================================================
' 11. GLOBAL EVENT COLLECTIONS
' =========================================================================
Public ButtonCollection As New Collection

' =========================================================================
' 12. MONTH NAMES
' =========================================================================
Public Const MONTH_NAMES_LIST As String = "Ocak,Şubat,Mart,Nisan,Mayıs,Haziran,Temmuz,Ağustos,Eylül,Ekim,Kasım,Aralık"
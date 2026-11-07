# =============================================================================
#  Report theme - classic blue (similar to the original look of these reports)
#  See default.psd1 for what every setting does.
# =============================================================================
@{
    Font            = "'Segoe UI', Arial, sans-serif"
    LabelFont       = "'Segoe UI', Arial, sans-serif"
    UppercaseLabels = $false
    ExcelFont       = 'Calibri'
    ExcelHeaderFont = 'Calibri'
    ExcelFontSize   = 11

    Primary         = '#305496'
    OnPrimary       = '#ffffff'
    Heading         = '#305496'
    Subtitle        = '#1f2328'
    Text            = '#1f2328'
    Muted           = '#6e7781'
    Page            = '#ffffff'
    Card            = '#ffffff'
    Line            = '#d0d7de'
    Rule            = '#d0d7de'
    RowStripe       = '#f6f8fa'

    StatusBad       = @('#842029', '#f8d7da', '#f1aeb5')
    StatusWarning   = @('#664d03', '#fff3cd', '#ffe69c')
    StatusInfo      = @('#084298', '#e7f1ff', '#b6d4fe')
    StatusGood      = @('#0f5132', '#d1e7dd', '#a3cfbb')
}

# =============================================================================
#  Report theme - the built-in look (warm neutrals, monospace labels)
#  https://github.com/nafthan/powershell-scripts
#
#  A theme sets the fonts and colours of the HTML reports, the emails and the
#  Excel workbooks. Copy this file, change what you like, then use it with:
#    Toolkit:     Theme = @{ File = 'themes\my-theme.psd1' } in config.psd1
#    Standalone:  .\Some-Report.ps1 -ThemeFile .\my-theme.psd1
#  Or create one from a Word/Excel template:  .\themes\New-ReportTheme.ps1
#
#  Leave out any setting to keep its default. Colours are hex, e.g. '#1f3a5f'.
# =============================================================================
@{
    # --- Labels ------------------------------------------------------------
    Organisation    = ''            # shown in the top bar and footer, e.g. 'Contoso Ltd'. Empty = hidden
    Classification  = 'Internal'    # the label at the top right, e.g. 'Confidential'. Empty = hidden
    LogoUrl         = ''            # optional https:// address of a logo for the top bar (HTML only, ~22px high)

    # --- Fonts -------------------------------------------------------------
    Font            = "'IBM Plex Sans', 'Segoe UI', Helvetica, Arial, sans-serif"   # HTML text, first installed font wins
    LabelFont       = "'IBM Plex Mono', Consolas, 'Courier New', monospace"        # HTML labels and table headings
    UppercaseLabels = $true         # small spaced capitals for labels; $false for normal case
    ExcelFont       = 'Segoe UI'    # Excel uses one font name, so pick one installed on Windows
    ExcelHeaderFont = 'Consolas'    # Excel header row and summary labels
    ExcelFontSize   = 10

    # --- Colours -----------------------------------------------------------
    Primary         = '#1c1b19'     # top bar and table header rows
    OnPrimary       = '#e8e4dc'     # text on Primary
    Accent          = ''            # optional stripe under the top bar and line under headings (e.g. a brand colour)
    Heading         = '#1c1b19'     # report title and section headings
    Subtitle        = '#45423d'     # intro text under the title
    Text            = '#1c1b19'     # table text
    Muted           = '#78736a'     # small labels and the footer
    Page            = '#f4f2ee'     # background around the report
    Card            = '#ffffff'     # report background
    Line            = '#e6e2da'     # row lines and borders
    Rule            = '#d6d1c8'     # lines around the details block
    RowStripe       = ''            # optional colour for every other row
    LabelFill       = ''            # optional background for the details block and Excel summary labels

    # --- Status colours: text, background, border --------------------------
    StatusBad       = @('#9b2c2c', '#f9ebe9', '#ecc9c4')   # Critical, High, Expired, Failed ...
    StatusWarning   = @('#875a0b', '#fbf3e1', '#ecd9ab')   # Warning, Medium, Expiring
    StatusInfo      = @('#2d5a86', '#eaf1f8', '#c7d7e8')   # Low
    StatusGood      = @('#2f6a47', '#e9f3ec', '#c4ddcc')   # OK, Done
}

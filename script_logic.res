# [Power Automate Desktop] - Automatización de Validación de Cobranza
# Propósito: Optimizar la verificación de estados de cuenta mediante RPA.
# Skill: Automatización Web + Manipulación de Excel.

# 1. Preparación de archivos de trabajo
Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess Path: $'''C:\\Users\\User\\Documents\\BookCELULARES.xlsx''' Visible: True ReadOnly: False Instance=> ExcelInstance
Excel.ReadFromExcel.ReadAllCells Instance: ExcelInstance GetCellContentsMode: Excel.GetCellContentsMode.TypedValues FirstLineIsHeader: True RangeValue=> ExcelData

Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess Path: $'''C:\\Users\\User\\Documents\\BookRESULTADOS.xlsx''' Visible: True ReadOnly: False Instance=> ExcelInstance2

# 2. Navegación al Portal Administrativo
WebAutomation.LaunchEdge.LaunchEdge Url: $'''https://www.sistema-interno-empresa.com/login''' WindowState: WebAutomation.BrowserWindowState.Normal BrowserInstance=> Browser

WebAutomation.PopulateTextField.PopulateTextFieldUsePhysicalKeyboard BrowserInstance: Browser Control: appmask['Input text \'username\''] Text: $'''USUARIO_ADMIN''' 
WebAutomation.PopulateTextField.PopulateTextFieldUsePhysicalKeyboard BrowserInstance: Browser Control: appmask['Input password \'password\''] Text: $'''********''' 
WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Input submit \'Iniciar Sesión\'']

# 3. Navegación a la sección de consulta
WAIT (WebAutomation.WaitForWebPageContent.WebPageToContainElement BrowserInstance: Browser Control: appmask['Anchor \'CONSULTA\''])
WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Anchor \'CONSULTA\'']
WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Anchor \'Estado de Cuenta\'']

Excel.GetFirstFreeRowOnColumn Instance: ExcelInstance2 Column: $'''A''' FirstFreeRowOnColumn=> FirstFreeRowOnColumn

# 4. Ciclo de procesamiento (Validación por registro)
LOOP FOREACH CurrentItem IN ExcelData
    # Ingresar dato de búsqueda
    WebAutomation.PopulateTextField.PopulateTextFieldUsePhysicalKeyboard BrowserInstance: Browser Control: appmask['Input text \'telefono_input\''] Text: CurrentItem
    WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Span \'Consultar\'']
    
    # Abrir detalle de cuenta
    WAIT (WebAutomation.WaitForWebPageContent.WebPageToContainElement BrowserInstance: Browser Control: appmask['Anchor \'verEdoCuenta\''])
    WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Anchor \'verEdoCuenta\'']
    
    # 5. Extracción y comparación de datos (Lógica de Negocio)
    # Se extrae el texto del estado de cuenta para validar la fecha de pago
    WebAutomation.ExtractData.ExtractSingleValue BrowserInstance: Browser ExtractionParameters: {[$'''html > body > table''', $'''Own Text''', $''''''] } TimeoutInSeconds: 60 ExtractedData=> DataFromWebPage
    
    ON ERROR
       # Manejo de excepciones en caso de registro no encontrado
    END
    
    # Si el pago coincide con la fecha objetivo, se registra en el reporte
    IF Contains(DataFromWebPage, $'''15/01/2026''', False) THEN
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance2 Value: CurrentItem Column: $'''A''' Row: FirstFreeRowOnColumn
        Variables.IncreaseVariable Value: FirstFreeRowOnColumn IncrementValue: 1
    END
    
    # Regresar para la siguiente consulta
    WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Anchor \'CONSULTA\'']
    WebAutomation.PressButton.PressButton BrowserInstance: Browser Control: appmask['Span \'Estado de Cuenta\'']
END

# 6. Notificación de fin de proceso
Display.ShowMessageDialog.ShowMessage Title: $'''Proceso Finalizado''' Message: $'''El robot ha terminado de procesar la lista de registros con éxito.''' Icon: Display.Icon.Information
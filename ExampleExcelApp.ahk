#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "ExcelLibrary\ExcelManager.ahk"
#Include "Util\Utils.ahk"
#Include "Util\OrObject.ahk"

ExampleExcelApp().Run()

/**
 * @internal
 * Aplicación de ejemplo que integra automatización de Excel.
 * 
 * Arquitectura modular:
 * - ExampleApp      -> Orquestador de alto nivel (eventos UI <-> servicios)
 * - ExampleWindow   -> Interfaz gráfica (presentación pura, sin lógica)
 * - ExcelService    -> Envuelve ExcelManager y notifica a la UI
 * 
 * @author bitasuperactive, DeepSeek (estructura)
 */
class ExampleExcelApp
{
    /** @type {ExampleWindow} */
    Window := unset
    /** @type {ExcelService} */
    Excel  := unset

    /**
     * Usa `Run()` para iniciar.
     */
    __New() {
        this.Window := ExampleWindow()
        this.Excel  := ExcelService(this.Window)
        this._WireEvents()
    }

    /**
     * Muestra la interfaz.
     */
    Run() => this.Window.Show()

    ; --- Composición de eventos entre UI y servicios -----------------------
    _WireEvents() {
        w := this.Window

        ; Excel
        w.OnExcelConnect     := (*)    => this.Excel.Connect()
        w.OnExcelDisconnect  := (*)    => this.Excel.Disconnect()
        w.OnExcelOpenFile    := (_, type) => this._OpenWorkbookFile(type)
        w.OnExcelReadSelect  := (_, name) => this.Excel.SelectReadWorkbook(name)
        w.OnExcelReadCell    := (_, row, col)  => this.Excel.ReadCell(row, col)
        w.OnExcelWriteSelect := (_, name) => this.Excel.SelectWriteWorkbook(name)
        w.OnExcelCreateTable := (*)    => this.Excel.CreateExampleTable()
        w.OnExcelReadTable   := (*)    => this.Excel.ReadTable()
        w.OnExcelValidate    := (*)    => this.Excel.ValidateHeaders()
        w.OnExcelReadHeaders := (*)    => this.Excel.ReadHeaders()
    }

    ; --- Handlers Excel ----------------------------------------------------
    _OpenWorkbookFile(type) {
        path := FileSelect(1,, "Selecciona un libro de Excel", "Libro de Excel (*.xlsx)")
        if (!path || !InStr(path, ".xlsx"))
            return
        Run(path)
        Sleep(1000)

        if (!this.Excel.IsConnected())
            this.Excel.Connect()
        this.Excel.RefreshWorkbookList()

        idx := this.Window.GetWorkbookItemCount()
        if (type = "read")
            this.Window.SelectReadWorkbookByIndex(idx)
        else
            this.Window.SelectWriteWorkbookByIndex(idx)
    }
}


; ============================================================================
; ExampleWindow - Interfaz de usuario (solo presentación)
; ============================================================================
class ExampleWindow
{
    ; --- Eventos expuestos a ExampleApp -----------------------------------
    OnExcelConnect     := (*) => 0
    OnExcelDisconnect  := (*) => 0
    OnExcelOpenFile    := (t) => 0
    OnExcelReadSelect  := (n) => 0
    OnExcelWriteSelect := (n) => 0
    OnExcelCreateTable := (*) => 0
    OnExcelReadTable   := (*) => 0
    OnExcelReadCell    := (r, c) => 0
    OnExcelValidate    := (*) => 0
    OnExcelReadHeaders := (*) => 0

    ; --- Estado interno ----------------------------------------------------
    /** @type {ExampleWindow} */
    _form    := unset
    /** @type {Map<String, Gui.Control>} */
    _ctrl    := Map()
    /** @type {Map<String, Microsoft.Office.Interop.Excel.Workbook>} */
    _wbItems := Map()

    __New() {
        this._BuildUI()
        this._WireInternalEvents()
    }

    Show() => this._form.Show()

    ; --- API pública para los servicios ------------------------------------
    GetWorkbookItemCount() => this._wbItems.Count

    SetExcelConnected(connected) {
        this._ctrl["ExcelConnectBtn"].Enabled    := !connected
        this._ctrl["ExcelDisconnectBtn"].Enabled := connected
        this._ctrl["ExcelGroup"].Text := connected ? "EXCEL ✅" : "EXCEL"

        if (!connected) {
            this._ctrl["ReadList"].Delete()
            this._ctrl["WriteList"].Delete()
            this._wbItems.Clear()
            this.SetReadReady(false)
            this.SetWriteReady(false)
        }
    }

    SetReadReady(ready) {
        this._ctrl["ReadTableBtn"].Enabled := ready
        this._ctrl["ValidateBtn"].Enabled  := ready
        this._ctrl["ColumnDropDown"].Enabled := ready
        this._UpdateCombinedState()
    }

    SetWriteReady(ready) {
        this._ctrl["CreateTableBtn"].Enabled := ready
        this._UpdateCombinedState()
    }

    UpdateWorkbookList(actualNames) {
        ; Eliminar libros que ya no están abiertos
        for name, idx in this._wbItems.Clone() {
            if (!Utils.ArrHasVal(actualNames, name)) {
                this._ctrl["ReadList"].Delete(idx)
                this._ctrl["WriteList"].Delete(idx)
                this._wbItems.Delete(name)
            }
        }
        ; Añadir nuevos
        for name in actualNames {
            if (this._wbItems.Has(name))
                continue
            this._ctrl["ReadList"].Add([name])
            this._ctrl["WriteList"].Add([name])
            this._wbItems.Set(name, this._wbItems.Count + 1)
        }
    }

    SelectReadWorkbookByIndex(idx)  => ControlChooseIndex(idx, this._ctrl["ReadList"])
    SelectWriteWorkbookByIndex(idx) => ControlChooseIndex(idx, this._ctrl["WriteList"])

    PopulateColumns(headers) {
        dd := this._ctrl["ColumnDropDown"]
        dd.Delete()
        if (headers.Length)
            dd.Add(headers)
        Hotkey("F1", (*) => MsgBox("Selecciona una columna a copiar.",, "0x40"), "On")
    }

    OnDropDownChange() {
        this._rowIndex := 1
        col := this._ctrl["ColumnDropDown"].Value ;int
        Hotkey("F1", (*) => this.CopyValueToClip(col), "On")
    }

    CopyValueToClip(col) {
        this._rowIndex += 1
        try {
            val := this.OnExcelReadCell(this._rowIndex, col)
            this._ctrl["ColumnRowIndex"].Text := this._rowIndex
            this._ctrl["ColumnValue"].Text := val
            A_Clipboard := val
        }
        catch {
            ;// Fuera del rango utilizado
            Hotkey("F1", (*) => this.CopyValueToClip(col), "off")
            MsgBox("Fin de rango.",, "0x40")
        }
    }

    ; --- Construcción de la UI ---------------------------------------------
    _BuildUI() {
        form := Gui("+Border +Caption -Resize", "ExampleGUI - Excel") ; 600x360
        form.SetFont("s10", "Segoe UI")
        this._form := form

        c := this._ctrl

        ; ---------------------------------------------------------------
        ; Fila 1 — Botones de Excel
        ; ---------------------------------------------------------------
        c["ExcelConnectBtn"]    := form.AddButton("x15  y15 w160 h32",           "Conectar Excel")
        c["ExcelDisconnectBtn"] := form.AddButton("x185 y15 w160 h32 Disabled",  "Desconectar Excel")

        ; ---------------------------------------------------------------
        ; Grupo EXCEL
        ; ---------------------------------------------------------------
        c["ExcelGroup"] := form.AddGroupBox("x15 y60 w570 h278", "EXCEL")

        ; Columna izquierda — Lectura
        form.AddText("x30  y90  w260 h20", "Libro de lectura")
        c["ReadList"]    := form.AddListBox("x30  y112 w260 h140")
        c["OpenReadBtn"] := form.AddButton("x256 y82 w34 h28", "📁")

        ; Columna derecha — Escritura
        form.AddText("x310 y90  w260 h20", "Libro de escritura")
        c["WriteList"]    := form.AddListBox("x310 y112 w260 h140")
        c["OpenWriteBtn"] := form.AddButton("x536 y82 w34 h28", "📁")

        ; Fila de acciones bajo las listas
        c["CreateTableBtn"] := form.AddButton("x30  y262 w180 h28 Disabled", "Crear tabla de ejemplo")
        c["ReadTableBtn"]   := form.AddButton("x220 y262 w120 h28 Disabled", "Leer tabla")
        c["ValidateBtn"]    := form.AddButton("x350 y262 w120 h28 Disabled", "Validar cabeceras")

        form.AddText("x30 y300 w65 h20", "Columna:")
        c["ColumnDropDown"] := form.AddDropDownList("x95 y300 w110 h95 Disabled")
        form.AddText("x225 y300 w70 h20", "[F1 to copy]")
        form.AddText("x310 y300 w30 h20", "Fila: ")
        c["ColumnRowIndex"]      := form.AddText("x335 y300 w10 h20", "0")
        form.AddText("x360 y300 w35 h20", "Valor: ")
        c["ColumnValue"]      := form.AddText("x400 y300 h20 w160", "")
    }

    _WireInternalEvents() {
        this._form.OnEvent("Close", (*) => ExitApp())

        c := this._ctrl
        c["ExcelConnectBtn"].OnEvent(   "Click",    (*) => this.OnExcelConnect())
        c["ExcelDisconnectBtn"].OnEvent("Click",    (*) => this.OnExcelDisconnect())
        c["OpenReadBtn"].OnEvent(       "Click",    (*) => this.OnExcelOpenFile("read"))
        c["OpenWriteBtn"].OnEvent(      "Click",    (*) => this.OnExcelOpenFile("write"))
        c["ReadList"].OnEvent(          "Change",   (ctrl, *) => this.OnExcelReadSelect(ctrl.Text))
        c["WriteList"].OnEvent(         "Change",   (ctrl, *) => this.OnExcelWriteSelect(ctrl.Text))
        c["CreateTableBtn"].OnEvent(    "Click",    (*) => this.OnExcelCreateTable())
        c["ReadTableBtn"].OnEvent(      "Click",    (*) => this.OnExcelReadTable())
        c["ValidateBtn"].OnEvent(       "Click",    (*) => this.OnExcelValidate())

        c["ColumnDropDown"].OnEvent(    "Focus",    (*) => this.PopulateColumns(this.OnExcelReadHeaders()))
        c["ColumnDropDown"].OnEvent(    "Change",   (*) => this.OnDropDownChange())
    }

    _UpdateCombinedState() {
        writeReady  := this._ctrl["CreateTableBtn"].Enabled
        enabled     := writeReady
    }
}


; ============================================================================
; ExcelService - Encapsula ExcelManager
; ============================================================================
class ExcelService
{
    /** @type {ExampleWindow} */
    _window     := unset
    /** @type {ExcelManager} */
    _manager    := unset
    _readReady  := false
    _writeReady := false

    __New(window) => this._window := window

    IsConnected()      => this.HasOwnProp("_manager") && this._manager != 0
    HasWriteWorkbook() => this._writeReady

    Connect() {
        if (this.IsConnected())
            return
        try {
            this._manager := ExcelManager(true)
        } catch Error as err {
            MsgBox("No se ha podido conectar con Excel:`n`n" err.Message, "Error", 16)
            return
        }

        ; --- Suscripción a eventos de Excel ---
        EEC := ExcelEventController
        EEC.OnEvent(EEC.ApplicationEventEnum.ANY_WORKBOOK_NEW,          (*) => this.RefreshWorkbookList())
        EEC.OnEvent(EEC.ApplicationEventEnum.ANY_WORKBOOK_OPEN,         (*) => SetTimer(() => this.RefreshWorkbookList(), -150))
        EEC.OnEvent(EEC.ApplicationEventEnum.ANY_WORKBOOK_AFTER_SAVE,   (*) => this.RefreshWorkbookList())
        EEC.OnEvent(EEC.ApplicationEventEnum.ANY_WORKBOOK_BEFORE_CLOSE, (*) => SetTimer(() => this.RefreshWorkbookList(), -150))
        EEC.OnEvent(EEC.ApplicationEventEnum.APPLICATON_TERMINATED,     (*) => this._OnClosed())

        this._window.SetExcelConnected(true)
        this.RefreshWorkbookList()
    }

    Disconnect() {
        if (!this.IsConnected())
            return
        try this._manager.Dispose()
        this._manager := unset
        this._readReady := false
        this._writeReady := false
        this._window.SetExcelConnected(false)
    }

    RefreshWorkbookList() {
        if (!this.IsConnected())
            return
        names := this._manager.GetAllOpenWorkbooksNames()
        this._window.UpdateWorkbookList(names)
    }

    SelectReadWorkbook(name) {
        if (!this.IsConnected() || !name)
            return
        try {
            this._manager.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.READ, name, false)
            this._readReady := true
            this._window.SetReadReady(true)

            headers := this._manager.ReadWorkbookAdapter.ReadRow(1)
            arr := []
            for prop in headers.OwnProps()
                arr.Push(prop)
            this._window.PopulateColumns(arr)
        } catch Error as err {
            MsgBox("Error conectando el libro de lectura:`n`n" err.Message, "Error", 16)
        }
    }

    SelectWriteWorkbook(name) {
        if (!this.IsConnected() || !name)
            return
        try {
            this._manager.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.WRITE, name, false)
            this._writeReady := true
            this._window.SetWriteReady(true)
        } catch Error as err {
            MsgBox("Error conectando el libro de escritura:`n`n" err.Message, "Error", 16)
        }
    }

    ReadCell(row, col) {
        row := this._manager.ReadWorkbookAdapter.ReadRow(row)
        for k, v in row.OwnProps()
            if A_Index = col
                return v
    }

    ReadHeaders() {
        headers := this._manager.ReadWorkbookAdapter.ReadRow(1)
        arr := []
        for prop in headers.OwnProps()
            arr.Push(prop)
        return arr
    }

    CreateExampleTable() {
        if (!this._writeReady)
            return
        objArray := []
        objArray.Push(OrObject(
            "Cuenta",    "Valor Cuenta 1",
            "Nombre",    "Valor Nombre 1",
            "Apellido",  "Valor Apellido 1",
            "Dirección", "Valor Dirección 1",
            "Teléfono",  "Valor Teléfono 1"))
        objArray.Push(OrObject(
            "Cuenta",    "Valor Cuenta 2",
            "Nombre",    "Valor Nombre 2",
            "Apellido",  "Valor Apellido 2",
            "Dirección", "Valor Dirección 2",
            "Teléfono",  "Valor Teléfono 2"))

        this._manager.WriteWorkbookAdapter.AppendTable(objArray)
        MsgBox("Tabla de ejemplo creada.", "Éxito", 64)
    }

    ReadTable() {
        if (!this._readReady)
            return
        adapter := this._manager.ReadWorkbookAdapter
        Loop adapter.GetRowCount() {
            row := adapter.ReadRow(A_Index)
            str := ""
            for k, v in row.OwnProps()
                str .= k ": " v "`n"
            MsgBox("[Fila " A_Index "]`n" str, "Contenido de la tabla")
        }
    }

    ValidateHeaders() {
        if (!this._readReady)
            return
        expected := ["APELLIDO", "NOMBRE", "DIRECCION", "TELEFONO", "CUENTA"]
        missing  := []
        ok := this._manager.ReadWorkbookAdapter.ValidateHeaders(expected, &missing)
        msg := ok ? "Todas las cabeceras requeridas están presentes."
                  : "Faltan cabeceras:`n" Utils.ArrayToString(missing)
        MsgBox(msg, ok ? "Éxito" : "Aviso", ok ? 64 : 48)
    }

    AppendTable(objArray) {
        if (!this._writeReady)
            return
        this._manager.WriteWorkbookAdapter.AppendTable(objArray)
    }

    _OnClosed() {
        if (!this.IsConnected())
            return
        this._manager := unset
        this._readReady := false
        this._writeReady := false
        this._window.SetExcelConnected(false)
        MsgBox("Excel se ha cerrado inesperadamente.", "Aviso", 48)
    }
}
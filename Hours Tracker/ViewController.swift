import Cocoa
import PythonKit

class ViewController: NSViewController {
    @IBOutlet weak var dateField: NSTextField!
    @IBOutlet weak var descriptionTextView: NSTextView!
    @IBOutlet weak var hoursField: NSTextField!
    @IBOutlet weak var hoursTypePopup: NSPopUpButton!
    @IBOutlet weak var usedHoursField: NSTextField!
    @IBOutlet weak var totalHoursBankedField: NSTextField!
    @IBOutlet weak var remainingHoursBankedField: NSTextField!
    @IBOutlet weak var promptLabel: NSTextField!
    
    var python: PythonObject!
    var getDateTop: NSWindow?
    
    override func viewDidLoad() {
        super.viewDidLoad()
        
        // Initialize Python
        python = PythonLibrary.useVersion(3)
        
        // Set window title
        view.window?.title = "SFNWA Hours Tracker"
        
        // Set the default values for hoursTypePopup
        hoursTypePopup.addItems(withTitles: ["Overtime", "Flex", "Used"])
        
        // Load the last window position if available
        if let position = UserDefaults.standard.string(forKey: "windowPosition") {
            view.window?.setFrameFromString(position)
        } else {
            // Center the window on the screen for the first time
            view.window?.center()
        }
    }
    
    @IBAction func onApplicationClose(_ sender: NSButton) {
        // Store the current window position before closing
        UserDefaults.standard.set(view.window?.frame.stringValue, forKey: "windowPosition")
        NSApplication.shared.terminate(sender)
    }
    
    @IBAction func getDate(_ sender: NSButton) {
        func setDate(_ selectedDate: Date) {
            let formatter = DateFormatter()
            formatter.dateFormat = "yyyy-MM-dd"
            dateField.stringValue = formatter.string(from: selectedDate)
            getDateTop?.close()
        }
        
        let calendarViewController = CalendarViewController()
        calendarViewController.onDateSelection = setDate(_:)
        getDateTop = NSWindow(contentViewController: calendarViewController)
        if let frame = view.window?.frame {
            let originX = frame.origin.x + frame.size.width
            let originY = frame.origin.y
            getDateTop?.setFrameOrigin(NSPoint(x: originX, y: originY))
        }
        getDateTop?.title = "Select Date"
        getDateTop?.makeKeyAndOrderFront(self)
    }
    
    @IBAction func calculateHours(_ sender: NSButton) {
        guard let hours = Double(hoursField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)) else {
            promptLabel.stringValue = "Invalid value for Hours Worked."
            promptLabel.textColor = .red
            return
        }
        
        let hoursType = hoursTypePopup.selectedItem?.title
        
        if hoursType == nil || !["Overtime", "Flex", "Used"].contains(hoursType!) {
            promptLabel.stringValue = "Invalid Hour Type."
            promptLabel.textColor = .red
            return
        }
        
        var totalHours = hours
        if hoursType == "Overtime" {
            totalHours *= 1.5
        }
        
        totalHoursBankedField.stringValue = String(format: "%.1f", totalHours)
        
        // Clear the prompt label
        promptLabel.stringValue = ""
        promptLabel.textColor = .labelColor
    }
    
    @IBAction func subtractHours(_ sender: NSButton) {
        guard let usedHours = Double(usedHoursField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)) else {
            promptLabel.stringValue = "Invalid value for Used Hours."
            promptLabel.textColor = .red
            return
        }
        
        guard let totalHoursBanked = Double(totalHoursBankedField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)) else {
            promptLabel.stringValue = "Calculate total hours banked first."
            promptLabel.textColor = .red
            return
        }
        
        let remainingHours = totalHoursBanked - usedHours
        remainingHoursBankedField.stringValue = String(format: "%.1f", remainingHours)
        
        // Clear the prompt label
        promptLabel.stringValue = ""
        promptLabel.textColor = .labelColor
    }
    
    @IBAction func saveToSpreadsheet(_ sender: NSButton) {
        let spreadsheetPath = "/Users/lisabains/downloads/hoursTracker.xlsx"
        print("Saving to spreadsheet:", spreadsheetPath)
        
        let selectedDate = dateField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)
        let description = descriptionTextView.string.trimmingCharacters(in: .whitespacesAndNewlines)
        let hours = hoursField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)
        let hoursType = hoursTypePopup.selectedItem?.title ?? ""
        let usedHours = usedHoursField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)
        
        if selectedDate.isEmpty || description.isEmpty || hours.isEmpty || hoursType.isEmpty || usedHours.isEmpty {
            promptLabel.stringValue = "All fields are required"
            promptLabel.textColor = .red
            return
        }
        
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "yyyy-MM-dd"
        guard let date = dateFormatter.date(from: selectedDate) else {
            promptLabel.stringValue = "Invalid date format. Please use YYYY-MM-DD."
            promptLabel.textColor = .red
            return
        }
        
        guard let hoursWorked = Double(hours) else {
            promptLabel.stringValue = "Invalid value for Hours Worked."
            promptLabel.textColor = .red
            return
        }
        
        guard let usedHoursValue = Double(usedHours) else {
            promptLabel.stringValue = "Invalid value for Used Hours."
            promptLabel.textColor = .red
            return
        }
        
        let remainingHoursBanked = Double(remainingHoursBankedField.stringValue) ?? 0
        
        if isSpreadsheetOpen() {
            let alert = NSAlert()
            alert.messageText = "Error"
            alert.informativeText = "Please close the spreadsheet first before saving."
            alert.runModal()
            return
        }
        
        python.run("""
        import openpyxl

        spreadsheet_path = "\(spreadsheetPath)"
        date = "\(selectedDate)"
        description = "\(description)"
        hours_worked = \(hoursWorked)
        hours_type = "\(hoursType)"
        used_hours = \(usedHoursValue)
        remaining_hours = \(remainingHoursBanked)

        try:
            wb = openpyxl.load_workbook(spreadsheet_path)
            sheet = wb.active
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(["Date", "Description", "Hours Worked", "Hours Type", "Used Hours", "Remaining Hours"])

        row = (date, description, hours_worked, hours_type, used_hours, remaining_hours)
        sheet.append(row)

        wb.save(spreadsheet_path)
        """)
        
        clearFields()
        totalHoursBankedField.stringValue = ""
    }
    
    func isSpreadsheetOpen() -> Bool {
        let processName: String
        if let bundleIdentifier = Bundle.main.bundleIdentifier {
            processName = bundleIdentifier
        } else {
            processName = "Python"
        }
        
        let task = Process()
        task.launchPath = "/bin/bash"
        task.arguments = ["-c", "pgrep -x \(processName)"]
        
        let pipe = Pipe()
        task.standardOutput = pipe
        task.launch()
        
        let data = pipe.fileHandleForReading.readDataToEndOfFile()
        let output = String(data: data, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines)
        
        return output?.isEmpty == false
    }
    
    @IBAction func loadSpreadsheet(_ sender: NSButton) {
        let spreadsheetPath = "/Users/lisabains/downloads/hoursTracker.xlsx"
        let workspace = NSWorkspace.shared
        workspace.open(URL(fileURLWithPath: spreadsheetPath))
    }
    
    @IBAction func closeSpreadsheet(_ sender: NSButton) {
        let processName: String
        if let bundleIdentifier = Bundle.main.bundleIdentifier {
            processName = bundleIdentifier
        } else {
            processName = "Python"
        }
        
        let task = Process()
        task.launchPath = "/bin/bash"
        task.arguments = ["-c", "pkill -x \(processName)"]
        task.launch()
    }
    
    @IBAction func clearFields(_ sender: NSButton) {
        dateField.stringValue = ""
        descriptionTextView.string = ""
        hoursField.stringValue = ""
        hoursTypePopup.selectItem(at: -1)
        usedHoursField.stringValue = ""
        totalHoursBankedField.stringValue = ""
        remainingHoursBankedField.stringValue = ""
        promptLabel.stringValue = ""
        promptLabel.textColor = .labelColor
    }
}

class CalendarViewController: NSViewController {
    @IBOutlet weak var calendarView: NSDatePicker!
    
    var onDateSelection: ((Date) -> Void)?
    
    override func viewDidLoad() {
        super.viewDidLoad()
        
        calendarView.datePickerStyle = .textFieldAndStepper
        calendarView.datePickerMode = .single
        calendarView.target = self
        calendarView.action = #selector(dateSelectionChanged)
        calendarView.dateValue = Date()
        calendarView.minDate = Date.distantPast
        calendarView.maxDate = Date()
    }
    
    @objc func dateSelectionChanged() {
        onDateSelection?(calendarView.dateValue)
    }
}

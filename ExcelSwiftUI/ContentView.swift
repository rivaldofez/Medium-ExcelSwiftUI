//
//  ContentView.swift
//  ExcelSwiftUI
//
//  Created by Rivaldo Fernandes on 13/07/23.
//

import SwiftUI
import xlsxwriter

struct ContentView: View {
    @State var documentItemsExport: [Any] = []
    @State private var showShareSheet: Bool = false
    
    
    var body: some View {
        VStack {
            ScrollView {
                ForEach(DummyData().contacts, id: \.name) { contact in
                    VStack {
                        Text(contact.name)
                            .font(.system(.title3).bold())
                        Text(contact.address)
                            .font(.system(.body))
                    }
                    .padding()
                }
            }
            
            Button("Export"){
                generateExcelFile(contacts: DummyData().contacts)
                self.showShareSheet.toggle()
            }
            
        }
        .padding()
        .sheet(isPresented: self.$showShareSheet) {
            ShareSheetView(items: self.$documentItemsExport)
        }
    }
    
    
    func generateExcelFile(contacts: [Contact]){
        let fileName = "DummyExcel.xlsx"
        let filePath = NSSearchPathForDirectoriesInDomains(.documentDirectory, .userDomainMask, true)[0].appending("/\(fileName)")
        
        let workbook = workbook_new(filePath)
        let worksheet = workbook_add_worksheet(workbook, nil)
        
        // Add style
        let format_header = workbook_add_format(workbook)
        format_set_bold(format_header)
        let format_1 = workbook_add_format(workbook)
        format_set_bg_color(format_1, 0xDDDDDD)
        
        //cell size
        let cell_width: Double = 50
        let cell_height: Double = 50
        var writingLine: UInt32 = 0
        
        //build header
        writingLine = 0
        let format = format_header
        format_set_bold(format)
        worksheet_write_string(worksheet, writingLine, 0, "No", format)
        worksheet_write_string(worksheet, writingLine, 1, "Image", format)
        worksheet_write_string(worksheet, writingLine, 2, "Name", format)
        worksheet_write_string(worksheet, writingLine, 3, "Gender", format)
        worksheet_write_string(worksheet, writingLine, 4, "Address", format)
        
        for contact in contacts {
            writingLine += 1
            
            worksheet_write_string(worksheet, writingLine, 0, "\(writingLine)", nil)
            worksheet_write_string(worksheet, writingLine, 2, contact.name, nil)
            worksheet_write_string(worksheet, writingLine, 3, contact.gender, nil)
            worksheet_write_string(worksheet, writingLine, 4, contact.address, nil)
            
        }
        
        for (index, contact) in contacts.enumerated() {
            let row = UInt32(index + 1)
            worksheet_set_row(worksheet, row, Double(cell_height), nil)
            let image = UIImage(systemName: contact.image)!
            var options = lxw_image_options()
            
            // Pixel size is Point size x image scale
            let imageScale = image.scale
            let uiimageSizeInPixel = (Double(image.size.width * imageScale), Double(image.size.height * imageScale))
            let scale = Helper.minRatio(left: (cell_width, cell_height),
                                        right: uiimageSizeInPixel )
            options.x_offset = 10
            options.y_offset = 1
            options.x_scale = scale
            options.y_scale = scale
            options.object_position = 1
            
            if let nsdata = image.jpegData(compressionQuality: 0.9) as NSData? {
                let buffer = Helper.getArrayOfBytesFromImage(imageData: nsdata)
                worksheet_insert_image_buffer_opt(worksheet, row, 1, buffer, buffer.count, &options)
            }
            
        }
        workbook_close(workbook)
        
        let docURL = URL(fileURLWithPath: filePath)
        self.documentItemsExport = [docURL as Any]
    }
}

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}

struct ShareSheetView: UIViewControllerRepresentable {
    typealias UIViewControllerType = UIActivityViewController
    
    @Binding var items: [Any]
    
    func makeUIViewController(context: UIViewControllerRepresentableContext<ShareSheetView>) -> UIActivityViewController {
        
        let controller = UIActivityViewController(activityItems: items, applicationActivities: nil)
        
        controller.completionWithItemsHandler = { (activityType: UIActivity.ActivityType?, completed: Bool, returnedItems: [Any]?, error: Error?) -> Void in
            if let url = items.first as? URL {
                try! FileManager.default.removeItem(at: url)
                items.removeAll()
            }
        }
        return controller
    }
    
    func updateUIViewController(_ uiViewController: UIActivityViewController, context: UIViewControllerRepresentableContext<ShareSheetView>) {
        
    }
}


class Helper {
    /// Update NSData to buffer for xlsxwriter
    static func getArrayOfBytesFromImage(imageData: NSData) -> [UInt8] {
        //Determine array size
        let count = imageData.length / MemoryLayout.size(ofValue: UInt8())
        //Create an array of the appropriate size
        var bytes = [UInt8](repeating: 0, count: count)
        //Copy image data as bytes into the array
        imageData.getBytes(&bytes, length:count * MemoryLayout.size(ofValue: UInt8()))

        return bytes
    }
    
    static func minRatio(left: (Double, Double), right: (Double, Double)) -> Double {
        min(left.0 / right.0, left.1 / right.1)
    }
    
}


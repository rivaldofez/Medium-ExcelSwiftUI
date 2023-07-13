//
//  Contact.swift
//  ExcelSwiftUI
//
//  Created by Rivaldo Fernandes on 13/07/23.
//

import SwiftUI

struct Contact {
    var name: String
    var gender: String
    var address: String
    var image: String
}

class DummyData {
    let contact1 = Contact(name: "Normal", gender: "Male", address: "Address 1", image: "person.fill")
    
    let contact2 = Contact(name: "Right", gender: "Male", address: "Address 2", image: "person.fill.turn.right")
    
    let contact3 = Contact(name: "Left", gender: "Female", address: "Address 3", image: "person.fill.turn.left")
    
    let contact4 = Contact(name: "Bottom", gender: "Female", address: "Address 4", image: "person.fill.turn.down")
    
    lazy var contacts: [Contact] =
    [ contact1, contact2, contact3, contact4 ]
}

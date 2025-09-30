#Requires AutoHotkey v2.0
MsgBox "Running Script."


::s1gn::
{
    ; Get individual components from A_Now
    currentYear := SubStr(A_Now, 1, 4)
    currentMonthNum := SubStr(A_Now, 5, 2)
    currentDay := SubStr(A_Now, 7, 2)

    ; Convert month number to month name
    monthNames := ["January","February","March","April","May","June","July","August","September","October","November","December"]
    currentMonth := monthNames[currentMonthNum]

    ; Build the formatted date
    currentDate := currentDay " of " currentMonth " " currentYear

    ; Send the multi-line text with dynamic date
    SendText "Good morning, Ken.`n`n"
    SendText "Please see the attached document for invoices for the orders delivered on the " currentDate ".`n`n"
    SendText "If you have any further requirements or questions, please do not hesitate to contact us.`n`n"
    SendText "Thank you and have a pleasant afternoon.`n`n"
    SendText "Regards,`n"
    SendText "CSR`n"
    SendText "Front Desk`n`n"
    SendText "Concord Aluminum Railings Limited`n"
    SendText "127 Corstate Avenue`n"
    SendText "Concord, Ontario`n"
    SendText "L4K 4Y2`n"
    SendText "PH: 905-669-0878  Ext. 221`n"
    SendText "Fax: 905-669-1268`n"
    SendText "www.concordaluminumrailings.com`n"
}

; Quote Email for 36" Standard Picket Railings
::q36s::
{
    ; Normal text first
    SendText "Hi -----,`n`n"
    SendText "As requested, please see attached "
    
    ; Bold "rough quote"
    Send("^b")
    SendText "rough quote"
    Send("^b")
    SendText " for a 36`"" " high picket railings on ------- Top in Black.`n`n"

    SendText "Based on your measurements provided I have the suggested "
    
    ; Bold "cutting measurements"
    Send("^b")
    SendText "cutting measurements"
    Send("^b")
    SendText " on the drawing.`n`n"

    SendText "Kindly review the attached quote and let me know if anything has been missed or requires adjustment.`n`n"

    SendText "To proceed with final pricing, we’ll need a "
    
    ; Bold "final drawing/sketch"
    Send("^b")
    SendText "final drawing/sketch"
    Send("^b")
    SendText " or lists showing inside post-to-post measurements ("
    SendText "cutting measurements"
    SendText ") and for stair sections: "
    
    ; Bold "stair angle"
    Send("^b")
    SendText "stair angle"
    Send("^b")
    SendText " and "
    
    ; Bold "sloping length"
    Send("^b")
    SendText "sloping length"
    Send("^b")
    SendText ".`n`nPlease refer to the attached sample layout for guidance. If needed, we can also provide cut-off posts for measurement purposes.`n`n"

    ; Bold "Important Notes:"
    Send("^b")
    SendText "Important Notes:`n"
    SendText "Maximum section length:"
    Send("^b")
    SendText " 6 feet (per engineering standards)`n"
    SendText "To proceed with your order, we require:`n"

    SendText "Written confirmation via email`n"

    ; Bold "50% deposit"
    Send("^b")
    SendText "50% deposit"
    SendText "`n"

    ; Bold "Payment methods:"
    SendText "Payment methods:"
    Send("^b")
    SendText " Visa/Mastercard (in person or by phone), cash, or cheque`n`n"

    ; Bold "Current lead time for picket railings"
    Send("^b")
    SendText "Current lead time for picket railings"
    Send("^b")
    SendText ": approx. "
    
    ; Bold "2 to 3 weeks"
    Send("^b")
    SendText "2 to 3 weeks"
    Send("^b")
    SendText "`n`nIf you have any questions or need further clarification, feel free to reach out.`n`n"

    ; Bold "Thanks!"
    Send("^b")
    SendText "Thanks!`n`n"
    Send("^b")

    ; Bold "Regards," and "Jason"
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    Send("^b")
    SendText "Regards,`n`nJason"
    Send("^b")
}

; Quote Email For 36" Glass Railings
::q36g::
{
    ; Normal text first
    SendText "Hi -----,`n`n"
    SendText "As requested, please see attached "
    
    ; Bold "rough quote"
    Send("^b")
    SendText "rough quote"
    Send("^b")
    SendText " for a 36`"" " high glass railings on ------- Top in Black.`n`n"

    SendText "Based on your measurements provided I have the suggested "
    
    ; Bold "cutting measurements"
    Send("^b")
    SendText "cutting measurements"
    Send("^b")
    SendText " on the drawing.`n`n"

    SendText "Kindly review the attached quote and let me know if anything has been missed or requires adjustment.`n`n"

    SendText "To proceed with final pricing, we’ll need a "
    
    ; Bold "final drawing/sketch"
    Send("^b")
    SendText "final drawing/sketch"
    Send("^b")
    SendText " or lists showing inside post-to-post measurements ("
    SendText "cutting measurements"
    SendText ") and for stair sections: "
    
    ; Bold "stair angle"
    Send("^b")
    SendText "stair angle"
    Send("^b")
    SendText " and "
    
    ; Bold "sloping length"
    Send("^b")
    SendText "sloping length"
    Send("^b")
    SendText ".`n`nPlease refer to the attached sample layout for guidance. If needed, we can also provide cut-off posts for measurement purposes.`n`n"

    ; Bold "Important Notes:"
    Send("^b")
    SendText "Important Notes:`n"
    SendText "Maximum section length:"
    Send("^b")
    SendText " 6 feet (per engineering standards)`n"
    SendText "To proceed with your order, we require:`n"

    SendText "Written confirmation via email`n"

    ; Bold "50% deposit"
    Send("^b")
    SendText "50% deposit"
    SendText "`n"

    ; Bold "Payment methods:"
    SendText "Payment methods:"
    Send("^b")
    SendText " Visa/Mastercard (in person or by phone), cash, or cheque`n`n"

    ; Bold "Current lead time for picket railings"
    Send("^b")
    SendText "Current lead time for picket railings"
    Send("^b")
    SendText ": approx. "
    
    ; Bold "3 to 4 weeks"
    Send("^b")
    SendText "3 to 4 weeks"
    Send("^b")
    SendText "`n`nIf you have any questions or need further clarification, feel free to reach out.`n`n"

    ; Bold "Thanks!"
    Send("^b")
    SendText "Thanks!`n`n"
    Send("^b")

    ; Bold "Regards," and "Jason"
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    Send("^b")
    SendText "Regards,`n`nJason"
    Send("^b")
}

; Quote Email For 42 Standard Picket Railings
::q42s::
{
    ; Normal text first
    SendText "Hi -----,`n`n"
    SendText "As requested, please see attached "
    
    ; Bold "rough quote"
    Send("^b")
    SendText "rough quote"
    Send("^b")
    SendText " for a 42`"" " high picket railings on ------- Top in Black.`n`n"

    SendText "Based on your measurements provided I have the suggested "
    
    ; Bold "cutting measurements"
    Send("^b")
    SendText "cutting measurements"
    Send("^b")
    SendText " on the drawing.`n`n"

    SendText "Kindly review the attached quote and let me know if anything has been missed or requires adjustment.`n`n"

    SendText "To proceed with final pricing, we’ll need a "
    
    ; Bold "final drawing/sketch"
    Send("^b")
    SendText "final drawing/sketch"
    Send("^b")
    SendText " or lists showing inside post-to-post measurements ("
    SendText "cutting measurements"
    SendText ") and for stair sections: "
    
    ; Bold "stair angle"
    Send("^b")
    SendText "stair angle"
    Send("^b")
    SendText " and "
    
    ; Bold "sloping length"
    Send("^b")
    SendText "sloping length"
    Send("^b")
    SendText ".`n`nPlease refer to the attached sample layout for guidance. If needed, we can also provide cut-off posts for measurement purposes.`n`n"

    ; Bold "Important Notes:"
    Send("^b")
    SendText "Important Notes:`n"
    SendText "Maximum section length:"
    Send("^b")
    SendText " 6 feet (per engineering standards)`n"
    SendText "To proceed with your order, we require:`n"

    SendText "Written confirmation via email`n"

    ; Bold "50% deposit"
    Send("^b")
    SendText "50% deposit"
    SendText "`n"

    ; Bold "Payment methods:"
    SendText "Payment methods:"
    Send("^b")
    SendText " Visa/Mastercard (in person or by phone), cash, or cheque`n`n"

    ; Bold "Current lead time for picket railings"
    Send("^b")
    SendText "Current lead time for picket railings"
    Send("^b")
    SendText ": approx. "
    
    ; Bold "2 to 3 weeks"
    Send("^b")
    SendText "2 to 3 weeks"
    Send("^b")
    SendText "`n`nIf you have any questions or need further clarification, feel free to reach out.`n`n"

    ; Bold "Thanks!"
    Send("^b")
    SendText "Thanks!`n`n"
    Send("^b")

    ; Bold "Regards," and "Jason"
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    Send("^b")
    SendText "Regards,`n`nJason"
    Send("^b")
}

; Quote Email For 42 Glass Railings
::q42g::
{
    ; Normal text first
    SendText "Hi -----,`n`n"
    SendText "As requested, please see attached "
    
    ; Bold "rough quote"
    Send("^b")
    SendText "rough quote"
    Send("^b")
    SendText " for a 42`"" " high glass railings on ------- Top in Black.`n`n"

    SendText "Based on your measurements provided I have the suggested "
    
    ; Bold "cutting measurements"
    Send("^b")
    SendText "cutting measurements"
    Send("^b")
    SendText " on the drawing.`n`n"

    SendText "Kindly review the attached quote and let me know if anything has been missed or requires adjustment.`n`n"

    SendText "To proceed with final pricing, we’ll need a "
    
    ; Bold "final drawing/sketch"
    Send("^b")
    SendText "final drawing/sketch"
    Send("^b")
    SendText " or lists showing inside post-to-post measurements ("
    SendText "cutting measurements"
    SendText ") and for stair sections: "
    
    ; Bold "stair angle"
    Send("^b")
    SendText "stair angle"
    Send("^b")
    SendText " and "
    
    ; Bold "sloping length"
    Send("^b")
    SendText "sloping length"
    Send("^b")
    SendText ".`n`nPlease refer to the attached sample layout for guidance. If needed, we can also provide cut-off posts for measurement purposes.`n`n"

    ; Bold "Important Notes:"
    Send("^b")
    SendText "Important Notes:`n"
    SendText "Maximum section length:"
    Send("^b")
    SendText " 6 feet (per engineering standards)`n"
    SendText "To proceed with your order, we require:`n"

    SendText "Written confirmation via email`n"

    ; Bold "50% deposit"
    Send("^b")
    SendText "50% deposit"
    SendText "`n"

    ; Bold "Payment methods:"
    SendText "Payment methods:"
    Send("^b")
    SendText " Visa/Mastercard (in person or by phone), cash, or cheque`n`n"

    ; Bold "Current lead time for picket railings"
    Send("^b")
    SendText "Current lead time for picket railings"
    Send("^b")
    SendText ": approx. "
    
    ; Bold "3 to 4 weeks"
    Send("^b")
    SendText "3 to 4 weeks"
    Send("^b")
    SendText "`n`nIf you have any questions or need further clarification, feel free to reach out.`n`n"

    ; Bold "Thanks!"
    Send("^b")
    SendText "Thanks!`n`n"
    Send("^b")

    ; Bold "Regards," and "Jason"
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    SendText ""
    Send("^b")
    SendText "Regards,`n`nJason"
    Send("^b")
}

; Word Order Status
::s1at::
{
    ; Normal text first
    SendText "Hi -----,`n`n"
    SendText "Good -------!`n`n"
    SendText "Good news! Your order will be ready for pick up -------------------------.`n`n"
    SendText "Please contact us when you plan to pick them up and to settle the amount. I have attached the invoice for your reference.`n`n"
    SendText "You’re welcome to give me a call to process the payment. We accept the following payment methods:`n"
    SendText "Visa/Mastercard (in person or by phone)`n"
    SendText "Cash`n"
    SendText "Cheque`n`n"

    ; Bolded section
    Send("^b") ; Turn bold on
    SendText "Pickup Hours:`n"
    SendText "Monday to Thursday:" 
    Send("^b")
    SendText "9:00 AM – 3:30 PM`n"
    Send("^b")
    SendText "Friday: "
    Send("^b")
    SendText "9:00 AM – 2:30 PM`n`n"
    Send("^b")
    SendText "Please avoid the following break times:`n"
    SendText "Morning Break: "
    Send("^b")
    SendText "9:45 – 10:00 AM`n"
    Send("^b")
    SendText "Lunch Break:"
    Send("^b")
    SendText "12:45 – 1:15 PM`n`n"
    Send("^b")
    SendText "Have a great day!!!`n`n"
    SendText "Regards,`n`n"
    SendText "Jason`n"
    Send("^b") ; Turn bold off just in case
}


;Closing Script
F2::{
    MsgBox "Script Closing"
    ExitApp
}

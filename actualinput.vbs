On Error Resume Next
Const PAGE_LOADED = 4

'Set objIE = CreateObject("InternetExplorer.Application")
'objIE.Visible = True

set shapp=createobject("shell.application")

For Each owin In shapp.Windows

set objIE = owin

err.clear

Next

On error goto 0

WScript.Sleep(300)

num = 0
On Error Resume Next
set sizes = objIE.Document.all.Item("size").length
On error goto 0

do while (num < 1000)

if sizes > 3 then
On Error Resume Next
objIE.Document.all.Item("size").SelectedIndex = "2"
On error goto 0
exit do

else
On Error Resume Next
objIE.Document.all.Item("size").SelectedIndex = "1"
On error goto 0
exit do

end if
WScript.Sleep(20)
num = num + 1
loop

On Error Resume Next
WScript.Sleep(300)
objIE.Document.all.Item("commit").click()
WScript.Sleep(250)
On error goto 0

objIE.Navigate("http://www.supremenewyork.com/checkout")

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0

On Error Resume Next
WScript.Sleep(100)

objIE.Document.all.Item("order[billing_name]").Value = "xxx-your name-xxx"
objIE.Document.all.Item("order[email]").Value = "xxx-your email-xxx"
objIE.Document.all.Item("order[tel]").Value = "xxx-your phone number-xxx"
objIE.Document.all.Item("order[billing_address]").Value = "xxx-your address-xxx"
objIE.Document.all.Item("order[billing_city]").Value = "xxx-yourcity-xxx"
objIE.Document.all.Item("order[billing_zip]").Value = "xxx-yourpostcode-xxx"
objIE.Document.all.Item("credit_card[cnb]").Value = "xxx-yourcreditcard number-xxx"
objIE.Document.all.Item("credit_card[vval]").Value = "xxx-yourcreditcard ccv-xxx"
objIE.Document.all.Item("order[billing_country]").Value = "xxx-yourcreditcard country-xxx"
objIE.Document.all.Item("credit_card[month]").Value = "xxx-yourcreditcard month-xxx"
objIE.Document.all.Item("credit_card[year]").Value = "xxx-yourcreditcard year-xxx"

On Error GoTo 0
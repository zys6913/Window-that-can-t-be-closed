const title = "Window that can't be closed"
const yourname = "I am a pig"
const question = "What kind of animal are you?"
const info = "You are not this animal"
const scend = "Haha, are you fooled~"
dim youranswer
do
youranswer = inputbox(question, title)
if youranswer <> yourname then msgbox info, vbinformation+vbokonly, title
loop until youranswer = yourname
msgbox scend, vbinformation+vbokonly, title

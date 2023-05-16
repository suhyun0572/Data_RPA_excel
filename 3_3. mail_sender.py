import win32com.client

def mail_sender(sending_file_list=[]):
    # HTML for the email
    before = '''<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8">
                <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
                o\:* {behavior:url(#default#VML);}
                w\:* {behavior:url(#default#VML);}
                .shape {behavior:url(#default#VML);}
                </style><![endif]--><style><!--
                /* Font Definitions */
                @font-face
                    {font-family:"Cambria Math";
                    panose-1:2 4 5 3 5 4 6 3 2 4;}
                @font-face
                    {font-family:Calibri;
                    panose-1:2 15 5 2 2 2 4 3 2 4;}
                @font-face
                    {font-family:"맑은 고딕";
                    panose-1:2 11 5 3 2 0 0 2 0 4;}
                @font-face
                    {font-family:"\@맑은 고딕";}
                /* Style Definitions */
                p.MsoNormal, li.MsoNormal, div.MsoNormal
                    {margin:0cm;
                    text-align:justify;
                    text-justify:inter-ideograph;
                    text-autospace:none;
                    word-break:break-hangul;
                    font-size:10.0pt;
                    font-family:"맑은 고딕";}
                span.EmailStyle17
                    {mso-style-type:personal-compose;
                    font-family:"맑은 고딕";
                    color:windowtext;}
                .MsoChpDefault
                    {mso-style-type:export-only;
                    font-family:"맑은 고딕";}
                .MsoPapDefault
                    {mso-style-type:export-only;
                    text-align:justify;
                    text-justify:inter-ideograph;
                    text-autospace:none;
                    word-break:break-hangul;}
                /* Page Definitions */
                @page WordSection1
                    {size:612.0pt 792.0pt;
                    margin:3.0cm 72.0pt 72.0pt 72.0pt;}
                div.WordSection1
                    {page:WordSection1;}
                --></style><!--[if gte mso 9]><xml>
                <o:shapedefaults v:ext="edit" spidmax="1026" />
                </xml><![endif]--><!--[if gte mso 9]><xml>
                <o:shapelayout v:ext="edit">
                <o:idmap v:ext="edit" data="1" />
                </o:shapelayout></xml><![endif]--></head><body lang=KO link="#0563C1" vlink="#954F72" style='word-wrap:break-word'><div class=WordSection1><p class=MsoNormal><span lang=DE>'''
    after = '''<o:p></o:p></span></p><p class=MsoNormal><span lang=DE><o:p>&nbsp;</o:p></span></p><p class=MsoNormal style='background:white'><span lang=DE style='font-size:12.0pt;font-family:"Arial",sans-serif;color:#222222'>&nbsp;<o:p></o:p></span></p><p class=MsoNormal style='background:white'><span lang=DE style='font-family:"Arial",sans-serif;color:#222222'>Mit freundlichen Gr&uuml;ßen / Best regards</span><span lang=DE style='font-size:12.0pt;font-family:"Arial",sans-serif;color:#222222'><o:p></o:p></span></p><p class=MsoNormal style='background:white'><b><span lang=DE style='font-family:"Arial",sans-serif;color:#222222'>Suhyun Lee</span></b><span lang=DE style='font-size:12.0pt;font-family:"Arial",sans-serif;color:#222222'><o:p></o:p></span></p><p class=MsoNormal style='background:white'><span lang=DE style='font-size:12.0pt;font-family:"Arial",sans-serif;color:#222222'>&nbsp;<o:p></o:p></span></p><p class=MsoNormal style='background:white'><span lang=DE style='font-size:12.0pt;font-family:"Arial",sans-serif;color:#222222'>'''
    
    new_Mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
    # email subject
    new_Mail.Subject = 'Automatic email for receiving schedules'
    # email contents
    new_Mail.HTMLBody = before+'Automatic email'+after
    # email to send
    new_Mail.To = 'test@test.com'
    # file attach
    if sending_file_list:
        for file in sending_file_list:
            new_Mail.Attachments.Add(file)
    
    # mail send
    if sending_file_list != []:
        new_Mail.Send()
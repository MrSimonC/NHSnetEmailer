__author__ = 'nbf1707'
#python 2.7

def SendNHSMail(fromEmail, Password, To, Subject, Body=""):
    import mechanize

    br = mechanize.Browser()
    #User-Agent Chrome 37
    br.addheaders = [('User-agent', 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2049.0 Safari/537.36')]
    br.set_handle_robots(False)

    #Logon
    urlLogOn = "https://web.nhs.net/CookieAuth.dll?GetLogon?curl=Z2Fportal&reason=0&formdir=5"
    br.open(urlLogOn)
    br.select_form(nr=0)
    br.form.set_all_readonly(False)
    br.form['curl']='Z2Fportal'
    br.form['flags']='0'
    br.form['forcedownlevel']='0'
    br.form['formdir']='5'
    br.form['partUsername']=fromEmail
    br.form['password']=Password
    br.form['SubmitCreds']='Log in'
    br.form['trusted']=['4']
    br.form['username']=fromEmail
    br.submit()

    #open Outlook url
    outlookurl = 'https://web.nhs.net/OWA/'
    br.open(outlookurl)
    #br._ua_handlers['_cookies'].cookiejar
    #get cookie info/userContext string for later
    cookieData = br._ua_handlers['_cookies'].cookiejar
    cookieData = str(cookieData)
    usercontext = cookieData[cookieData.find("UserContext=")+12:cookieData.find(". for web.nhs.net")+1]

    #open new Mail
    newMailurl = "https://web.nhs.net/OWA/?ae=Item&t=IPM.Note&a=New"
    br.open(newMailurl)

    #send new Mail
    import urllib
    postEmailToURL = "https://web.nhs.net/OWA/?ae=PreFormAction&t=IPM.Note&a=Send"
    params = {
        'hidpnst' : '',
        'txtto' : To,
        'txtcc' : '',
        'txtbcc' : '',
        'txtsbj' : Subject,
        'txtbdy' : Body,
        'hidid' : '',
        'hidchk' : '',
        'hidunrslrcp' : '0',
        'hidcmdpst' : "snd",
        'hidrmrcp' : '',
        'hidaddrcptype' : '',
        'hidaddrcp' : '',
        'hidss' : '',
        'hidmsgimp' : '1',
        'hidrcptid' : '',
        'hidrw' : '',
        'hidpid' : "EditMessage",
        'hidcanary' : usercontext
    }
    data = urllib.urlencode(params)
    br.open(postEmailToURL, data)
    print br.response().read()


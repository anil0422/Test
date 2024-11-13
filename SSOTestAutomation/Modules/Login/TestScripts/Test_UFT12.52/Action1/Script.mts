SystemUtil.Run "firefox.exe","https://sso-itg.ext.hpe.com"
Browser("Welcome - HPE Software").Page("Welcome - HPE Software").WebButton("My Software Support sign-in").Click
Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").WebEdit("username").Set "auto_uft92@hpe.com"
Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").WebEdit("password").SetSecure "570b489196c01effbd6c49347b4d238cbd499bb7e759"
Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").Link("Sign in").Click
Browser("Welcome - HPE Software").Page("My Support Home - HPE").Link("Sign Out").Click

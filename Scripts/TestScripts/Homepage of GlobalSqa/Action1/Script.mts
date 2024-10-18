Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Contact Us").Click
Browser("Home - GlobalSQA").Page("Contact Us - GlobalSQA").WebEdit("comment_name").Set "makshi"
Browser("Home - GlobalSQA").Page("Contact Us - GlobalSQA").WebEdit("email").Set "makshi@gmail.com"
Browser("Home - GlobalSQA").Page("Contact Us - GlobalSQA").WebEdit("comment_name").Set "makshi"
Browser("Home - GlobalSQA").Page("Contact Us - GlobalSQA").WebEdit("subject").Set "testing"
Browser("Home - GlobalSQA").Page("Contact Us - GlobalSQA").WebEdit("comment").Set "just for testing purposes. Please ignore this."

' Wait for the frame to load
Wait(3)

' Use Descriptive Programming to identify the frame and the checkbox
Set recaptchaFrame = Browser("title:=Home - GlobalSQA").Page("title:=Contact Us - GlobalSQA").Frame("title:=reCAPTCHA")
Set recaptchaCheckbox = recaptchaFrame.WebElement("html tag:=SPAN", "html id:=recaptcha-anchor", "class:=recaptcha-checkbox goog-inline-block recaptcha-checkbox-unchecked rc-anchor-checkbox recaptcha-checkbox-hover recaptcha-checkbox-focused recaptcha-checkbox-clearOutline")

' Click the checkbox
recaptchaCheckbox.Click

' Click the Send button
Browser("title:=Contact Us - GlobalSQA").Page("title:=Contact Us - GlobalSQA").WebButton("html tag:=BUTTON", "innertext:=Send").Click

' Close the Browser
Browser("title:=Contact Us - GlobalSQA").Close


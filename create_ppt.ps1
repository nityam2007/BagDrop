# PowerShell Script to Create BagDrop PPT
# Swiss Style: Bold, Yellow Accent, Black/White
# Saves to Desktop to avoid Unicode path issues

Add-Type -AssemblyName System.Drawing

$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

$presentation = $pptApp.Presentations.Add()

# Define colors (BGR format for PowerPoint)
$colorBlack = 0
$colorWhite = 16777215
$colorYellow = 55295  # BGR: FFD700 -> 00D7FF in BGR = 55295
$colorDarkGray = 1973790

$ppLayoutBlank = 12

# ============== SLIDE 1: TITLE ==============
$slide1 = $presentation.Slides.Add(1, $ppLayoutBlank)
$slide1.FollowMasterBackground = 0
$slide1.Background.Fill.ForeColor.RGB = $colorWhite

$bar1 = $slide1.Shapes.AddShape(1, 0, 180, 150, 200)
$bar1.Fill.ForeColor.RGB = $colorYellow
$bar1.Line.Visible = 0

$title1 = $slide1.Shapes.AddTextbox(1, 100, 150, 600, 100)
$title1.TextFrame.TextRange.Text = "BAGDROP"
$title1.TextFrame.TextRange.Font.Size = 72
$title1.TextFrame.TextRange.Font.Bold = -1
$title1.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title1.TextFrame.TextRange.Font.Name = "Arial Black"

$sub1 = $slide1.Shapes.AddTextbox(1, 100, 260, 600, 50)
$sub1.TextFrame.TextRange.Text = "SMART LUGGAGE STORAGE & DELIVERY"
$sub1.TextFrame.TextRange.Font.Size = 24
$sub1.TextFrame.TextRange.Font.Bold = -1
$sub1.TextFrame.TextRange.Font.Color.RGB = $colorDarkGray
$sub1.TextFrame.TextRange.Font.Name = "Arial"

$tag1 = $slide1.Shapes.AddTextbox(1, 100, 320, 600, 40)
$tag1.TextFrame.TextRange.Text = "Secure. Contactless. Affordable."
$tag1.TextFrame.TextRange.Font.Size = 18
$tag1.TextFrame.TextRange.Font.Color.RGB = $colorYellow
$tag1.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 2: PROBLEM ==============
$slide2 = $presentation.Slides.Add(2, $ppLayoutBlank)
$slide2.FollowMasterBackground = 0
$slide2.Background.Fill.ForeColor.RGB = $colorWhite

$bar2 = $slide2.Shapes.AddShape(1, 500, 0, 220, 150)
$bar2.Fill.ForeColor.RGB = $colorYellow
$bar2.Line.Visible = 0

$title2 = $slide2.Shapes.AddTextbox(1, 50, 50, 400, 60)
$title2.TextFrame.TextRange.Text = "THE PROBLEM"
$title2.TextFrame.TextRange.Font.Size = 48
$title2.TextFrame.TextRange.Font.Bold = -1
$title2.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title2.TextFrame.TextRange.Font.Name = "Arial Black"

$prob = $slide2.Shapes.AddTextbox(1, 50, 150, 620, 300)
$prob.TextFrame.TextRange.Text = "Travelers check out early but flights/trains are later`n`nCarrying luggage all day is inconvenient & tiring`n`nCloakrooms are crowded, expensive, or unavailable`n`nForgotten luggage creates stress and financial loss"
$prob.TextFrame.TextRange.Font.Size = 22
$prob.TextFrame.TextRange.Font.Color.RGB = $colorDarkGray
$prob.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 3: SOLUTION ==============
$slide3 = $presentation.Slides.Add(3, $ppLayoutBlank)
$slide3.FollowMasterBackground = 0
$slide3.Background.Fill.ForeColor.RGB = $colorBlack

$bar3 = $slide3.Shapes.AddShape(1, 0, 300, 720, 15)
$bar3.Fill.ForeColor.RGB = $colorYellow
$bar3.Line.Visible = 0

$title3 = $slide3.Shapes.AddTextbox(1, 50, 80, 620, 80)
$title3.TextFrame.TextRange.Text = "OUR SOLUTION"
$title3.TextFrame.TextRange.Font.Size = 48
$title3.TextFrame.TextRange.Font.Bold = -1
$title3.TextFrame.TextRange.Font.Color.RGB = $colorYellow
$title3.TextFrame.TextRange.Font.Name = "Arial Black"

$sol = $slide3.Shapes.AddTextbox(1, 50, 180, 620, 250)
$sol.TextFrame.TextRange.Text = "BAGDROP - Store Bags Anywhere, Anytime`n`nPartner with local shops for secure storage`nSafe, contactless access via mobile app`nOTP, NFC tokens, UPI & Card payments`nSelf-service kiosks at transit hubs"
$sol.TextFrame.TextRange.Font.Size = 22
$sol.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$sol.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 4: VALUE PROPOSITION ==============
$slide4 = $presentation.Slides.Add(4, $ppLayoutBlank)
$slide4.FollowMasterBackground = 0
$slide4.Background.Fill.ForeColor.RGB = $colorYellow

$title4 = $slide4.Shapes.AddTextbox(1, 50, 50, 620, 80)
$title4.TextFrame.TextRange.Text = "WHY BAGDROP?"
$title4.TextFrame.TextRange.Font.Size = 56
$title4.TextFrame.TextRange.Font.Bold = -1
$title4.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title4.TextFrame.TextRange.Font.Name = "Arial Black"

$val = $slide4.Shapes.AddTextbox(1, 50, 160, 620, 300)
$val.TextFrame.TextRange.Text = "Rs.50 for 6 hours - Affordable pricing`n`nMulti-payment: UPI, Card, NFC`n`nSmart OTP & Token access system`n`nDelivery support for forgotten luggage`n`nInter-city rapid porter arrangements"
$val.TextFrame.TextRange.Font.Size = 24
$val.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$val.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 5: TARGET MARKET ==============
$slide5 = $presentation.Slides.Add(5, $ppLayoutBlank)
$slide5.FollowMasterBackground = 0
$slide5.Background.Fill.ForeColor.RGB = $colorWhite

$bar5 = $slide5.Shapes.AddShape(1, 0, 0, 40, 540)
$bar5.Fill.ForeColor.RGB = $colorYellow
$bar5.Line.Visible = 0

$title5 = $slide5.Shapes.AddTextbox(1, 60, 50, 620, 80)
$title5.TextFrame.TextRange.Text = "TARGET AUDIENCE"
$title5.TextFrame.TextRange.Font.Size = 48
$title5.TextFrame.TextRange.Font.Bold = -1
$title5.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title5.TextFrame.TextRange.Font.Name = "Arial Black"

$aud = $slide5.Shapes.AddTextbox(1, 60, 160, 620, 300)
$aud.TextFrame.TextRange.Text = "-> Tourists & Backpackers`n`n-> Train & Bus Travelers`n`n-> Students & Daily Commuters`n`n-> Business Travelers`n`n-> Anyone with forgotten luggage needs"
$aud.TextFrame.TextRange.Font.Size = 26
$aud.TextFrame.TextRange.Font.Color.RGB = $colorDarkGray
$aud.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 6: FEATURES ==============
$slide6 = $presentation.Slides.Add(6, $ppLayoutBlank)
$slide6.FollowMasterBackground = 0
$slide6.Background.Fill.ForeColor.RGB = $colorWhite

$title6 = $slide6.Shapes.AddTextbox(1, 50, 30, 620, 60)
$title6.TextFrame.TextRange.Text = "KEY FEATURES"
$title6.TextFrame.TextRange.Font.Size = 44
$title6.TextFrame.TextRange.Font.Bold = -1
$title6.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title6.TextFrame.TextRange.Font.Name = "Arial Black"

$f1 = $slide6.Shapes.AddShape(1, 50, 110, 200, 100)
$f1.Fill.ForeColor.RGB = $colorYellow
$f1.Line.Visible = 0
$f1.TextFrame.TextRange.Text = "APP BOOKING`n& MAP VIEW"
$f1.TextFrame.TextRange.Font.Size = 16
$f1.TextFrame.TextRange.Font.Bold = -1
$f1.TextFrame.TextRange.Font.Name = "Arial Black"

$f2 = $slide6.Shapes.AddShape(1, 260, 110, 200, 100)
$f2.Fill.ForeColor.RGB = $colorBlack
$f2.Line.Visible = 0
$f2.TextFrame.TextRange.Text = "SELF-SERVICE`nKIOSK"
$f2.TextFrame.TextRange.Font.Size = 16
$f2.TextFrame.TextRange.Font.Bold = -1
$f2.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$f2.TextFrame.TextRange.Font.Name = "Arial Black"

$f3 = $slide6.Shapes.AddShape(1, 470, 110, 200, 100)
$f3.Fill.ForeColor.RGB = $colorYellow
$f3.Line.Visible = 0
$f3.TextFrame.TextRange.Text = "NFC & QR`nTOKEN"
$f3.TextFrame.TextRange.Font.Size = 16
$f3.TextFrame.TextRange.Font.Bold = -1
$f3.TextFrame.TextRange.Font.Name = "Arial Black"

$f4 = $slide6.Shapes.AddShape(1, 50, 220, 200, 100)
$f4.Fill.ForeColor.RGB = $colorBlack
$f4.Line.Visible = 0
$f4.TextFrame.TextRange.Text = "REAL-TIME`nNOTIFICATIONS"
$f4.TextFrame.TextRange.Font.Size = 16
$f4.TextFrame.TextRange.Font.Bold = -1
$f4.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$f4.TextFrame.TextRange.Font.Name = "Arial Black"

$f5 = $slide6.Shapes.AddShape(1, 260, 220, 200, 100)
$f5.Fill.ForeColor.RGB = $colorYellow
$f5.Line.Visible = 0
$f5.TextFrame.TextRange.Text = "DELIVERY`nSUPPORT"
$f5.TextFrame.TextRange.Font.Size = 16
$f5.TextFrame.TextRange.Font.Bold = -1
$f5.TextFrame.TextRange.Font.Name = "Arial Black"

$f6 = $slide6.Shapes.AddShape(1, 470, 220, 200, 100)
$f6.Fill.ForeColor.RGB = $colorBlack
$f6.Line.Visible = 0
$f6.TextFrame.TextRange.Text = "24/7`nSUPPORT"
$f6.TextFrame.TextRange.Font.Size = 16
$f6.TextFrame.TextRange.Font.Bold = -1
$f6.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$f6.TextFrame.TextRange.Font.Name = "Arial Black"

# ============== SLIDE 7: HOW IT WORKS ==============
$slide7 = $presentation.Slides.Add(7, $ppLayoutBlank)
$slide7.FollowMasterBackground = 0
$slide7.Background.Fill.ForeColor.RGB = $colorBlack

$title7 = $slide7.Shapes.AddTextbox(1, 50, 30, 620, 60)
$title7.TextFrame.TextRange.Text = "HOW IT WORKS"
$title7.TextFrame.TextRange.Font.Size = 44
$title7.TextFrame.TextRange.Font.Bold = -1
$title7.TextFrame.TextRange.Font.Color.RGB = $colorYellow
$title7.TextFrame.TextRange.Font.Name = "Arial Black"

$steps = $slide7.Shapes.AddTextbox(1, 50, 110, 620, 350)
$steps.TextFrame.TextRange.Text = "01  FIND`n    Locate nearby BagDrop location on app`n`n02  BOOK`n    Reserve your time slot instantly`n`n03  DROP`n    Store luggage & receive OTP/NFC token`n`n04  RETRIEVE`n    Pick up using your secure token`n`n05  FORGOT?`n    Request delivery via support line"
$steps.TextFrame.TextRange.Font.Size = 18
$steps.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$steps.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 8: TECHNOLOGY ==============
$slide8 = $presentation.Slides.Add(8, $ppLayoutBlank)
$slide8.FollowMasterBackground = 0
$slide8.Background.Fill.ForeColor.RGB = $colorWhite

$bar8 = $slide8.Shapes.AddShape(1, 600, 0, 120, 540)
$bar8.Fill.ForeColor.RGB = $colorYellow
$bar8.Line.Visible = 0

$title8 = $slide8.Shapes.AddTextbox(1, 50, 50, 520, 60)
$title8.TextFrame.TextRange.Text = "TECH & PAYMENTS"
$title8.TextFrame.TextRange.Font.Size = 44
$title8.TextFrame.TextRange.Font.Bold = -1
$title8.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title8.TextFrame.TextRange.Font.Name = "Arial Black"

$tech = $slide8.Shapes.AddTextbox(1, 50, 140, 520, 300)
$tech.TextFrame.TextRange.Text = "PAYMENTS`nUPI / Card / NFC Contactless`n`nSECURITY`nOTP & Token Verification`nEnd-to-end encrypted`n`nPLATFORM`nMobile App (iOS & Android)`nSelf-service Kiosk UI`nAdmin Dashboard"
$tech.TextFrame.TextRange.Font.Size = 18
$tech.TextFrame.TextRange.Font.Color.RGB = $colorDarkGray
$tech.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 9: MANAGEMENT ==============
$slide9 = $presentation.Slides.Add(9, $ppLayoutBlank)
$slide9.FollowMasterBackground = 0
$slide9.Background.Fill.ForeColor.RGB = $colorYellow

$title9 = $slide9.Shapes.AddTextbox(1, 50, 30, 620, 60)
$title9.TextFrame.TextRange.Text = "TEAM STRUCTURE"
$title9.TextFrame.TextRange.Font.Size = 44
$title9.TextFrame.TextRange.Font.Bold = -1
$title9.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title9.TextFrame.TextRange.Font.Name = "Arial Black"

$roles = $slide9.Shapes.AddTextbox(1, 50, 110, 620, 350)
$roles.TextFrame.TextRange.Text = "CEO / FOUNDER`nOverall strategy & vision`n`nCTO`nApp & platform technology`n`nOPERATIONS MANAGER`nOnboarding & quality control`n`nLOGISTICS COORDINATOR`nDelivery partner management`n`nCUSTOMER SUPPORT`nPhone & app support team"
$roles.TextFrame.TextRange.Font.Size = 18
$roles.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$roles.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 10: MARKETING ==============
$slide10 = $presentation.Slides.Add(10, $ppLayoutBlank)
$slide10.FollowMasterBackground = 0
$slide10.Background.Fill.ForeColor.RGB = $colorWhite

$title10 = $slide10.Shapes.AddTextbox(1, 50, 50, 620, 60)
$title10.TextFrame.TextRange.Text = "GO-TO-MARKET"
$title10.TextFrame.TextRange.Font.Size = 48
$title10.TextFrame.TextRange.Font.Bold = -1
$title10.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title10.TextFrame.TextRange.Font.Name = "Arial Black"

$m1 = $slide10.Shapes.AddShape(1, 50, 130, 300, 80)
$m1.Fill.ForeColor.RGB = $colorYellow
$m1.Line.Visible = 0
$m1.TextFrame.TextRange.Text = "SOCIAL MEDIA ADS`nTargeting travelers"
$m1.TextFrame.TextRange.Font.Size = 14
$m1.TextFrame.TextRange.Font.Bold = -1
$m1.TextFrame.TextRange.Font.Name = "Arial Black"

$m2 = $slide10.Shapes.AddShape(1, 360, 130, 300, 80)
$m2.Fill.ForeColor.RGB = $colorBlack
$m2.Line.Visible = 0
$m2.TextFrame.TextRange.Text = "INFLUENCER COLLABS`nTravel content creators"
$m2.TextFrame.TextRange.Font.Size = 14
$m2.TextFrame.TextRange.Font.Bold = -1
$m2.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$m2.TextFrame.TextRange.Font.Name = "Arial Black"

$m3 = $slide10.Shapes.AddShape(1, 50, 220, 300, 80)
$m3.Fill.ForeColor.RGB = $colorBlack
$m3.Line.Visible = 0
$m3.TextFrame.TextRange.Text = "TRANSIT HUB FLYERS`nOn-ground presence"
$m3.TextFrame.TextRange.Font.Size = 14
$m3.TextFrame.TextRange.Font.Bold = -1
$m3.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$m3.TextFrame.TextRange.Font.Name = "Arial Black"

$m4 = $slide10.Shapes.AddShape(1, 360, 220, 300, 80)
$m4.Fill.ForeColor.RGB = $colorYellow
$m4.Line.Visible = 0
$m4.TextFrame.TextRange.Text = "REFERRAL BONUSES`nUsers & shop partners"
$m4.TextFrame.TextRange.Font.Size = 14
$m4.TextFrame.TextRange.Font.Bold = -1
$m4.TextFrame.TextRange.Font.Name = "Arial Black"

# ============== SLIDE 11: FINANCIALS ==============
$slide11 = $presentation.Slides.Add(11, $ppLayoutBlank)
$slide11.FollowMasterBackground = 0
$slide11.Background.Fill.ForeColor.RGB = $colorBlack

$title11 = $slide11.Shapes.AddTextbox(1, 50, 30, 620, 60)
$title11.TextFrame.TextRange.Text = "REVENUE MODEL"
$title11.TextFrame.TextRange.Font.Size = 44
$title11.TextFrame.TextRange.Font.Bold = -1
$title11.TextFrame.TextRange.Font.Color.RGB = $colorYellow
$title11.TextFrame.TextRange.Font.Name = "Arial Black"

$rev = $slide11.Shapes.AddTextbox(1, 50, 100, 620, 120)
$rev.TextFrame.TextRange.Text = "REVENUE STREAMS`nStorage fees - Rs.50 per booking`nDelivery arrangement fees`nNFC premium token fees`nKiosk advertisement revenue"
$rev.TextFrame.TextRange.Font.Size = 18
$rev.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$rev.TextFrame.TextRange.Font.Name = "Arial"

$big = $slide11.Shapes.AddTextbox(1, 50, 240, 620, 100)
$big.TextFrame.TextRange.Text = "Rs.7,50,000"
$big.TextFrame.TextRange.Font.Size = 72
$big.TextFrame.TextRange.Font.Bold = -1
$big.TextFrame.TextRange.Font.Color.RGB = $colorYellow
$big.TextFrame.TextRange.Font.Name = "Arial Black"

$bigSub = $slide11.Shapes.AddTextbox(1, 50, 340, 620, 40)
$bigSub.TextFrame.TextRange.Text = "ESTIMATED MONTHLY EARNINGS"
$bigSub.TextFrame.TextRange.Font.Size = 16
$bigSub.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$bigSub.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 12: EXIT STRATEGY ==============
$slide12 = $presentation.Slides.Add(12, $ppLayoutBlank)
$slide12.FollowMasterBackground = 0
$slide12.Background.Fill.ForeColor.RGB = $colorWhite

$bar12 = $slide12.Shapes.AddShape(1, 0, 400, 300, 140)
$bar12.Fill.ForeColor.RGB = $colorYellow
$bar12.Line.Visible = 0

$title12 = $slide12.Shapes.AddTextbox(1, 50, 50, 620, 60)
$title12.TextFrame.TextRange.Text = "EXIT STRATEGY"
$title12.TextFrame.TextRange.Font.Size = 48
$title12.TextFrame.TextRange.Font.Bold = -1
$title12.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title12.TextFrame.TextRange.Font.Name = "Arial Black"

$exit = $slide12.Shapes.AddTextbox(1, 50, 140, 620, 250)
$exit.TextFrame.TextRange.Text = "-> Acquisition by travel/logistics company`n`n-> Merge with major courier platform`n`n-> Technology IP licensing/sale`n`n-> Expansion to metro & international markets"
$exit.TextFrame.TextRange.Font.Size = 24
$exit.TextFrame.TextRange.Font.Color.RGB = $colorDarkGray
$exit.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 13: ROADMAP ==============
$slide13 = $presentation.Slides.Add(13, $ppLayoutBlank)
$slide13.FollowMasterBackground = 0
$slide13.Background.Fill.ForeColor.RGB = $colorYellow

$title13 = $slide13.Shapes.AddTextbox(1, 50, 50, 620, 60)
$title13.TextFrame.TextRange.Text = "FUTURE ROADMAP"
$title13.TextFrame.TextRange.Font.Size = 48
$title13.TextFrame.TextRange.Font.Bold = -1
$title13.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$title13.TextFrame.TextRange.Font.Name = "Arial Black"

$t1 = $slide13.Shapes.AddShape(1, 50, 140, 150, 80)
$t1.Fill.ForeColor.RGB = $colorBlack
$t1.Line.Visible = 0
$t1.TextFrame.TextRange.Text = "Q1`n50+ CITIES"
$t1.TextFrame.TextRange.Font.Size = 14
$t1.TextFrame.TextRange.Font.Bold = -1
$t1.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$t1.TextFrame.TextRange.Font.Name = "Arial Black"
$t1.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$t2 = $slide13.Shapes.AddShape(1, 210, 140, 150, 80)
$t2.Fill.ForeColor.RGB = $colorBlack
$t2.Line.Visible = 0
$t2.TextFrame.TextRange.Text = "Q2`nAIRPORTS"
$t2.TextFrame.TextRange.Font.Size = 14
$t2.TextFrame.TextRange.Font.Bold = -1
$t2.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$t2.TextFrame.TextRange.Font.Name = "Arial Black"
$t2.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$t3 = $slide13.Shapes.AddShape(1, 370, 140, 150, 80)
$t3.Fill.ForeColor.RGB = $colorBlack
$t3.Line.Visible = 0
$t3.TextFrame.TextRange.Text = "Q3`nSUBSCRIPTIONS"
$t3.TextFrame.TextRange.Font.Size = 14
$t3.TextFrame.TextRange.Font.Bold = -1
$t3.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$t3.TextFrame.TextRange.Font.Name = "Arial Black"
$t3.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$t4 = $slide13.Shapes.AddShape(1, 530, 140, 150, 80)
$t4.Fill.ForeColor.RGB = $colorBlack
$t4.Line.Visible = 0
$t4.TextFrame.TextRange.Text = "Q4`nSMART LOCKERS"
$t4.TextFrame.TextRange.Font.Size = 12
$t4.TextFrame.TextRange.Font.Bold = -1
$t4.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$t4.TextFrame.TextRange.Font.Name = "Arial Black"
$t4.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$roadNote = $slide13.Shapes.AddTextbox(1, 50, 260, 620, 100)
$roadNote.TextFrame.TextRange.Text = "VISION: Become India's #1 luggage storage network with smart locker integrations across all major transit hubs."
$roadNote.TextFrame.TextRange.Font.Size = 18
$roadNote.TextFrame.TextRange.Font.Color.RGB = $colorBlack
$roadNote.TextFrame.TextRange.Font.Name = "Arial"

# ============== SLIDE 14: THANK YOU ==============
$slide14 = $presentation.Slides.Add(14, $ppLayoutBlank)
$slide14.FollowMasterBackground = 0
$slide14.Background.Fill.ForeColor.RGB = $colorBlack

$thanks = $slide14.Shapes.AddTextbox(1, 50, 120, 620, 100)
$thanks.TextFrame.TextRange.Text = "THANK YOU"
$thanks.TextFrame.TextRange.Font.Size = 72
$thanks.TextFrame.TextRange.Font.Bold = -1
$thanks.TextFrame.TextRange.Font.Color.RGB = $colorYellow
$thanks.TextFrame.TextRange.Font.Name = "Arial Black"
$thanks.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$qna = $slide14.Shapes.AddTextbox(1, 50, 230, 620, 50)
$qna.TextFrame.TextRange.Text = "QUESTIONS & DISCUSSION"
$qna.TextFrame.TextRange.Font.Size = 24
$qna.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$qna.TextFrame.TextRange.Font.Name = "Arial"
$qna.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$contact = $slide14.Shapes.AddTextbox(1, 50, 320, 620, 80)
$contact.TextFrame.TextRange.Text = "[Your Email] | [Your Phone]`nwww.bagdrop.in"
$contact.TextFrame.TextRange.Font.Size = 18
$contact.TextFrame.TextRange.Font.Color.RGB = $colorWhite
$contact.TextFrame.TextRange.Font.Name = "Arial"
$contact.TextFrame.TextRange.ParagraphFormat.Alignment = 2

# Save to Desktop (avoiding Unicode path issues)
$savePath = [Environment]::GetFolderPath("Desktop") + "\BagDrop_Presentation.pptx"
$presentation.SaveAs($savePath)

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  SUCCESS! Presentation Created!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Saved to: $savePath" -ForegroundColor Cyan
Write-Host "Total Slides: 14" -ForegroundColor White
Write-Host ""
Write-Host "Design: Swiss Style (Bold, Yellow/Black/White)" -ForegroundColor White

                               #Date--11.08.2010

#**************************************PDU-1************************************
#PDU tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
#PDU1-Scroll to Branch level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 3).click

#PDU1 Settings Tab
$ie.frame(:index, 3).frame(:index, 3).image(:id, 'imgCtrlSettings').click

#PDU1 Settings Parameters
#PDU User Assigned Label
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'label').set('MPH Rack PDU1')
#PDU Asset Tag01
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'assetTag1').set('')
#PDU Asset Tag02
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'assetTag2').set('')
#Neutral Over Current Alarm Threshold
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'ecNeutralThrshldOverAlarm').set('80')
#Neutral Over Current Warning Threshold
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'ecNeutralThrshldOverWarn').set('75')
#Over Current Alarm Threshold
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'ecThresholdHiAlmLN').set('80')
#Over Current Warning Threshold
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'ecThresholdHiWrnLN').set('75')
#Low  Current Alarm  Threshold
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlApsSettingForm').text_field(:name, 'ecThresholdLoAlmLN').set('0')
 
#PDU1 Commands Tab    
$ie.frame(:index, 3).frame(:index, 3).image(:id, 'imgCtrlCommands').click
#PDU Accumulated Energy Reset button
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).button(:name, 'energyControl').click
#Generic Test Event
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).button(:name, 'testEvent').click

#PDU1-Branch1 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
#PDU1-Branch1 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 3).click

#PDU1-Branch1 Settings
#Branch User Assigned Label
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlRemSettingForm').text_field(:name, 'label').set('Branch 1')
#Branch Asset Tag01
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlRemSettingForm').text_field(:name, 'assetTag1').set('')
#Branch Asset Tag02
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlRemSettingForm').text_field(:name, 'assetTag2').set('')

#PDU1-Branch1-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click


#PDU1-Branch1-Receptacle1
#Receptacle User Assigned Label
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlReceptacleSettingForm').text_field(:id, 'label').set('Receptacle 1')
#Receptacle Asset Tag01
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlReceptacleSettingForm').text_field(:id, 'assetTag1').set('')
#Receptacle Asset Tag02
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).form(:name, 'rpcControlReceptacleSettingForm').text_field(:id, 'assetTag2').set('')

#Control Buttons
#Control-Edit
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).button(:name, 'Edit').click
#Control-Save
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).button(:name, 'Submit')
#Control-Reset
$ie.frame(:index, 3).frame(:index, 3).frame(:index, 1).button(:value, 'Reset')

#PDU Home
$ie.frame(:index, 3).frame(:index, 3).image(:index, 33).click
#Branch1-PDU1-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
#Receptacle-Branch1-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 29).click

#**************************************************************************************#



                                 #Date--12.08.2010
#****************************************PDU-2*************************************#
#PDU Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 10).click
#PDU2-Scroll to Branch level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click

#PDU2-Branch1 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
#PDU2-Branch1 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click

#PDU2-Branch1-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 6).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 10).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 18).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 22).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 30).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 34).click

#Branch1-PDU2-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
#Receptacle-Branch1-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 38).click

#***************************************PDU-3**********************************#
#PDU Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 18).click
#PDU3-Scroll to Branch level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 19).click

#PDU3-Branch1 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
#PDU3-Branch2 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
#PDU3-Branch3 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click

#PDU3-Branch1 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 3).click
#PDU3-Branch2 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 6).click
#PDU3-Branch3 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 9).click

#PDU3-Branch1-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click

#PDU3-Branch2-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click

#PDU3-Branch3-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click
   
#Branch-PDU3-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
#Receptacle-Branch1-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 29).click
#Receptacle-Branch2-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 29).click 
#Receptacle-Branch3-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 29).click 

#*******************************PDU-4***********************************#
#PDU Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click
#PDU4-Scroll to Branch level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 27).click


#PDU4-Branch1 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
#PDU4-Branch2 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
#PDU4-Branch3 Tab
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click

#PDU4-Branch1 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 3).click
#PDU4-Branch2 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 6).click
#PDU4-Branch3 Scroll to Receptacle Level
$ie.frame(:index, 3).frame(:index, 3).image(:index, 9).click

#PDU4-Branch1-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click

#PDU4-Branch2-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click

#PDU3-Branch3-Receptacle Tab(1 to 9)
$ie.frame(:index, 3).frame(:index, 3).image(:index, 2).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 5).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 8).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 14).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 17).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 20).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 23).click
$ie.frame(:index, 3).frame(:index, 3).image(:index, 26).click
   
#Branch-PDU4-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 11).click
#Receptacle-Branch1-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 38).click
#Receptacle-Branch2-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 38).click
#Receptacle-Branch3-Scroll back
$ie.frame(:index, 3).frame(:index, 3).image(:index, 38).click


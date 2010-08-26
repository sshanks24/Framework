                                  #Date 18-08-2010
								  
#**************************************Sensors************************************
#Click on Monitor Sensors link
$IE0.frame(:index, 3).frame(:index, 2).link(:text, 'Sensors').click

#Click on Status button
$IE0.frame(:index, 3).frame(:index, 3).image(:id, 'imgStatus').click

#Click on events button
$IE0.frame(:index, 3).frame(:index, 3).image(:id, 'imgEvents').click

#Click on Setting button
$IE0.frame(:index, 3).frame(:index, 3).image(:id, 'imgSettings').click

#Click on Ratings button
$IE0.frame(:index, 3).frame(:index, 3).image(:id, 'imgInfo').click

#click on Sensors ID Sort
$IE0.frame(:index, 3).frame(:index, 3).link(:text, 'ID').click

#click on Sensors Type Sort
$IE0.frame(:index, 3).frame(:index, 3).link(:text, 'Type').click

#click on Sensors Lablel Sort
$IE0.frame(:index, 3).frame(:index, 3).link(:text, 'Label').click

#click on Sensor ID link for ex :-4 -1-1
$IE0.frame(:index, 3).frame(:index, 3).link(:text, '4 -  1-1').click
$IE0.frame(:index, 3).frame(:index, 3).link(:text, '4 -  1-2').click
$IE0.frame(:index, 3).frame(:index, 3).link(:text, '4 -  1-3').click
$IE0.frame(:index, 3).frame(:index, 3).link(:text, '4 -  2-1').click
$IE0.frame(:index, 3).frame(:index, 3).link(:text, '4 -  2-2').click






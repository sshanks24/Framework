#**************************************Device-Explorer************************************#
#Device Explorer Link
$ie.frame(:index, 3).frame(:index, 2).link(:text, 'Device Explorer').click

#Label
$ie.frame(:index, 3).frame(:index, 3).form(:index, 1).text_field(:id, 'searchBox').set('')

#Device Explorer --Headers
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle User Assigned Label').click
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Id').click
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'O').click
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'C').click
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'S').click

#************************************PDU-1***************************************#
#Receptacle label(PDU1-Branch 1)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:1.1.1_0.1.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:1.1.2_0.2.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:1.1.3_0.3.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:1.1.4_0.4.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:1.1.5_0.5.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:1.1.6_0.6.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:1.1.7_0.7.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:1.1.8_0.8.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:1.1.9_0.9.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '1-1-9').click

#************************************PDU-2***************************************#

#Receptacle label(PDU2-Branch 1)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:2.1.1_0.37.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:2.1.2_0.38.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:2.1.3_0.39.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:2.1.4_0.40.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:2.1.5_0.41.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:2.1.6_0.42.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:2.1.7_0.43.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:2.1.8_0.44.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:2.1.9_0.45.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '2-1-9').click


#************************************PDU-3***************************************#
#Receptacle label(PDU3-Branch 1)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:3.1.1_0.73.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:3.1.2_0.74.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:3.1.3_0.75.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:3.1.4_0.76.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:3.1.5_0.77.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:3.1.6_0.78.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:3.1.7_0.79.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:3.1.8_0.80.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:3.1.9_0.81.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-1-9').click

#Receptacle label(PDU3-Branch 2)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:3.2.1_0.82.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:3.2.2_0.83.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:3.2.3_0.84.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:3.2.4_0.85.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:3.2.5_0.86.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:3.2.6_0.87.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:3.2.7_0.88.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:3.2.8_0.89.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:3.2.9_0.90.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-2-9').click

#Receptacle label(PDU3-Branch 3)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:3.3.1_0.91.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:3.3.2_0.92.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:3.3.3_0.93.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:3.3.4_0.94.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:3.3.5_0.95.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:3.3.6_0.96.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:3.3.7_0.97.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:3.3.8_0.98.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:3.3.9_0.99.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '3-3-9').click


#************************************PDU-4***************************************#
#Receptacle label(PDU4-Branch 1)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:4.1.1_0.109.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:4.1.2_0.110.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:4.1.3_0.111.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:4.1.4_0.112.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:4.1.5_0.113.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:4.1.6_0.114.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:4.1.7_0.115.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:4.1.8_0.116.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:4.1.9_0.117.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-1-9').click

#Receptacle label(PDU4-Branch 2)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:4.2.1_0.118.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:4.2.2_0.119.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:4.2.3_0.120.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:4.2.4_0.121.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:4.2.5_0.122.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:4.2.6_0.123.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:4.2.7_0.124.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:4.2.8_0.125.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:4.2.9_0.126.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-2-9').click

#Receptacle label(PDU4-Branch 3)
$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 1').click 
href = "javascript:rpcSelectCtx('std:4.3.1_0.127.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-1').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 2').click 
href = "javascript:rpcSelectCtx('std:4.3.2_0.128.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-2').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 3').click 
href = "javascript:rpcSelectCtx('std:4.3.3_0.129.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-3').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 4').click 
href = "javascript:rpcSelectCtx('std:4.3.4_0.130.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-4').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 5').click 
href = "javascript:rpcSelectCtx('std:4.3.5_0.131.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-5').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 6').click 
href = "javascript:rpcSelectCtx('std:4.3.6_0.132.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-6').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 7').click 
href = "javascript:rpcSelectCtx('std:4.3.7_0.133.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-7').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 8').click 
href = "javascript:rpcSelectCtx('std:4.3.8_0.134.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-8').click

$ie.frame(:index, 3).frame(:index, 3).link(:text, 'Receptacle 9').click 
href = "javascript:rpcSelectCtx('std:4.3.9_0.135.0');"
#Id
$ie.frame(:index, 3).frame(:index, 3).link(:text, '4-3-9').click
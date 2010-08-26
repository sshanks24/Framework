
19 July 2010
1) Add chng_ip3.rb  - yet another flavor for changing spreadsheet ip addresses
This is a refactoring of cng_ip2.rb



24 June 2010
1) Create this README file
2) Delete Utility_script.txt as it's contents were moved into this file.
2) Remove the version number from the utility script.
3) Add notes/instructions for all of the tools

4) chng_ip2.rb is a prototype used for changing the ip address in all of the driver
scripts by collecting the file names directly from the driver folder as opposed
to getting them from the controller spreadsheet.

5) chng_drvr_ver.rb is a proptotype that is used for changing the version number
of the driver.rb and driver.xls files. A user input determines if .rb, .xls, or
both will be updated. The method used increments the version number by 1.





18 May 2010

Add two lines in Utility_script2.rb
puts "Enter the Controller spreadsheet number: example - 44:"
ss1_no = gets.chomp!
Add one more parameter at line 33 in the array  'ss1_no'

Now utility file name is Utility_script3.rb


11 May 2010

Removed the following line from Utility2.rb

ss_name = 'C:/lmg_test/ruby/07.05.2010/driver/agent_table_info22.xls'


Use chng_ip if you just need to change the IP address.



07 May 2010
1) replace ip = ip.lstrip.rstrip with   ip.strip'-------------- at line 22
2) replace u_name = u_name.lstrip.rstrip with  u_name.strip'--- at line 26
3) replace pas_wd = pas_wd.lstrip.rstrip with  pas_wd.strip'----at line 29
4) replace all gsub methods with sub

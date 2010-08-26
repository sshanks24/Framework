=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*

cfg_snmp_traps7.rb

*Test_Case_Number*
 700.120.20.110-Positve and Negative Test Case

*Description*
  Validate the SNMP Traps Positive and Negative Information configuration

     - Identical Network Name and Community Name
     - Network Name and Community Name Boundary Test

     - Empty Network Name and Community Name
*Variable_Definitions*
    s = test start time
    f = test finish time
    e = test elapsed time
    roe = row number in controller spreadsheet
    excel = nested array that contains an instance of excel and driver parameters
    ss = spreadsheet
    wb = workbook
    ws = worksheet
    dvr_ss = driver spreadsheet
    rows = number of rows in spreadsheet to execute
    site = url/ip address of card being tested
    name = user name for login
    pswd = password for login

=end

#Launch the SNMP Access Negative ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now

require 'generic'
require 'watir/process' 

def traps_input(row,ws)
      # Read the input data from the spreadsheet into a multidimensional array
      # Where each sub-array represents one row of data
      #     1) network name     - 'k'
      #     2) Traps Port      -  'l'
      #     3) community name   - 'm'
	  #     4) Heart Beat       - 'n' 
      #     4) save flag        - 'o'  # ** Not used
      #     5) input row        - 'p'  # ** Not used
      #     6) popup expected   - 'af' # ** Not in the arrray currently **
      #
      input1 = ws.Range("k#{row}:p#{row+19}")['Value']
    end
	
def traps_fill(g,input1)
   #
   # Write data to all 20 rows of the traps table
   input1.each do|input|
      p input # show each row of input data
      i = input[5].to_i
	  p i
      g.trap_addr(i).set input[0].to_s
      g.trap_port(i).set input[1].to_s.delete(".0")
	  g.trap_com(i).set input[2].to_s
      g.trap_hb(i).set input[3]
  end
end	
	
def row_fill(g,ws,strt_row)

  for i in 1..20
    #puts "Value of i is #{i}"
    # puts "#{g.trap_addr(i).value}"
    ws.Range("bc#{strt_row}")['Value'] = g.trap_addr(i).value
    puts "#{g.trap_port(i).value}"
	ws.Range("bd#{strt_row}")['Value'] = g.trap_port(i).value
	ws.Range("be#{strt_row}")['Value'] = g.trap_com(i).value
    puts "#{g.trap_com(i).value}"
	ws.Range("bf#{strt_row}")['Value'] = g.checkbox(g.trap_hb(i))
	puts "#{g.checkbox(g.trap_hb(i))}"
    strt_row = strt_row + 1  
  end
end

begin 
  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  #Open up a new excel instance and save it with the timestamp.
  #Open the IE browser
  #Collect the support page information and save it in the time stamped spreadsheet.
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  rows = excel[1][1] 
  $ie.speed = :zippy
  #Navigate to the 'Configure’ tab
  g.config.click
  $ie.maximize  
  #Click the SNMP Traps link on the left side of widow
  #Login if not called from controller
  g.logn_chk(g.traps,excel[1])

  #Clear All SNMP Traps textboxes
  g.edit.click
  for i in 1..20
  g.trap_clr(i).click
  g.trap_hb(i).clear
  end
  g.save.click
  
  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    sleep 5
    Watir::Waiter.wait_until(5) { g.edit.exists?}
    g.edit.click   
	
	input1 = traps_input(row,ws)
    p input1 # show array with 20 rows of input data (for debug)
    p'-----------------'
	
	# Write data to all 20 rows of the access table
    traps_fill(g,input1)
	row += 19
	
    # Write SNMP Traps textboxes
    save_flag = ws.Range("o#{row}")['Value'].to_s
    pop = ws.Range("af#{row}")['Value'].to_s 
    unless pop == "no"
      ws.Range("bk#{row}")['Value'] = g.invChar($ie,pop,nil)
    else
      
      puts "Save Flag = #{save_flag}"
      if save_flag == "S"
        g.save.click
        #g.jsClick('Windows Internet Explorer', 'OK')
      end
    end

if save_flag == "S"
  #read SNMP Traps textboxes value
      sleep 3
      Watir::Waiter.wait_until(5) { g.edit.exists?}
      g.edit.click
      strt_row = row - 19    
      end_row = row
      puts "Start row = #{strt_row}"
      puts "End row = #{end_row}"
      row_fill(g,ws,strt_row)
      g.save.click
	  
  end
    wb.Save
  end

 f = Time.now  #finish time
#Capture error if any in the script  
rescue Exception => e
  f = Time.now  #finish time 
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s

ensure #this section is executed even if script goes in error
  if(error_present == nil)
    # If roe > 0, script is called from controller
    # If roe = 0, script is being ran independently
    #Close and save the spreadsheet and thes web browser.
    g.tear_down_d(excel[0],s,f,roe)
    if roe == 0
      $ie.close
    end
  else 
    puts" There were errors in the script"
    status = "script in error"
    wb.save
    wb.close
  end
end
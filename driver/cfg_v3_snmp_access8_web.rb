=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  snmp_access

*Test_Case_Number*
  700.190.20.110

*Description*
  Validate the SNMP Access Configuration
     - Positive
     - Negative
     - Boundary
    
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

#Launch the Configure SNMP Access ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'
require 'watir/process' 

def access_input(row,ws)
      # Read the input data from the spreadsheet into a multidimensional array
      # Where each sub-array represents one row of data
      #     1) network name     - 'k'
      #     2) access type      - 'l'
      #     3) community name   - 'm'
      #     4) save flag        - 'n'  # ** Not used
      #     5) input row        - 'o'  # ** Not used
      #     6) popup expected   - 'af' # ** Not in the arrray currently **
      #
      input1 = ws.Range("k#{row}:o#{row+19}")['Value']
    end

def access_fill(g,input1)
   #
   # Write data to all 20 rows of the access table
   input1.each do|input|
      p input # show each row of input data
      i = input[4].to_i
      g.access_addr(i).set input[0]
      if input[1].to_i == 1
        g.access_type((i),'1').set
      else
        g.access_type((i),'0').set
      end
      g.access_com(i).set input[2].to_s
  end
end

def row_fill(g,ws,strt_row)
  for i in 1..20
    puts "Value of i is #{i}"
    puts "#{g.access_addr(i).value}"
    ws.Range("bc#{strt_row}")['Value'] = g.access_addr(i).value
    puts "#{g.access_type(i,'1').checked?}"
    if g.access_type(i,'1').checked? == true
      ws.Range("bd#{strt_row}")['Value'] = "1"
    elsif g.access_type(i,'0').checked? == true
      ws.Range("bd#{strt_row}")['Value'] = "0"
    end
    puts "#{g.access_com(i).value}"
    ws.Range("be#{strt_row}")['Value'] = g.access_com(i).value
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
  #Navigate to the 'Configure' tab
  g.config.click
  $ie.maximize  
  #Click the Configure SNMP Access link on the left side of window
  #Login if not called from controller
  g.logn_chk(g.v1access,excel[1])
  
  #Clear All SNMP Access textboxes
  g.edit.click
  for i in 1..20
    g.access_clr(i).click
  end
  g.save.click


  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    sleep 5
    Watir::Waiter.wait_until(5) { g.edit.exists?}
    g.edit.click

    input1 = access_input(row,ws)
    p'----------------'
    p input1 # show array with 20 rows of input data
    p'-----------------'

    # Write data to all 20 rows of the access table
    access_fill(g,input1)

    row += 19
    p "at row #{row}" # which row are we at?

    save_flag = ws.Range("n#{row}")['Value'].to_s
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
      #read SNMP Access textboxes value
      sleep 3
      Watir::Waiter.wait_until(5) { g.edit.exists?}
      g.edit.click
      strt_row = row - 19    
      puts "Start row = #{strt_row}"
      puts "End row = #{row}"
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
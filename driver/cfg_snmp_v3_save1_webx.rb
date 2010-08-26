=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  manag_protocol_info for GXT/NX devices

*Test_Case_Number*
  700.090.20.110

*Description*
  Validate the Management Protocol Configuration Information
     - Enable / Disable SNMP Agent Check Box
     - Popup handle with Reset Ok button
     - Popup handle with Reset Cancel button

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

#Launch the Management Protocol ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'
require 'watir'

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
  #Navigate to the Configure tab
  g.config.click
  #Click the Management Protocl link on the left side of window
  #Login if not called from controller
  g.logn_chk(  g.mgtprot,excel[1])
  
  #select snmp v3 protocol
  g.edit.click
  g.snmp_v3.set
  g.save.click_no_wait

  g.snmpv3.click

  ##Clear all rows of any contents insnmp v3 table
  g.edit.click

  for i  in 1..20
    #puts i
    g.v3_clear(i).click
  end

  number = 1 #every user number

  for row in 1..rows
    row +=1 # add 1, execution starts at drvr_ss row 2
    puts " Executing -  Test step #{ws.Range("g#{row}")['Value'].to_i}"

    row_user = ws.Range("k#{row}:z#{row}")['Value']
    p "**********************************************************************"
    p row_user
    p "**********************************************************************"
    #Every 20 row start from user 1
    if number > 20
      number = 1
    end
    sleep 1

    Watir::Waiter.wait_until(500) { g.edit.exists?}

    #################write every user info############
    #1write enable user or not
    puts "1"
    if   row_user[0][0] == 'set' then g.v3_enable(number).set else g.v3_enable(number).clear end
    #2write user name
    puts "2"
    if    row_user[0][1].class == Float
      g.v3_user(number).set(row_user[0][1].floor.to_s)
    else
      g.v3_user(number).set( row_user[0][1].to_s)
    end
    #3write authentication
    puts "3"
    g.v3_auth(number, row_user[0][2].to_i).set
    #4write authentication secret
    puts "4"
    if  row_user[0][3].class == Float
      g.v3_auth_secret(number).set(row_user[0][3].floor.to_s)
    else
      g.v3_auth_secret(number).set(row_user[0][3].to_s)
    end
    #5write privacy
    puts "5"
    g.v3_privacy(number,row_user[0][4].to_i).set
    #6write Privacy secret
    puts "6"
    if  row_user[0][5].class == Float
      g.v3_privacy_secret(number).set(row_user[0][5].floor.to_s)
    else
      g.v3_privacy_secret(number).set(row_user[0][5].to_s)
    end
    #7write access read
    puts "7"
    if row_user[0][6] == 'set' then g.v3_acc_read(number).set else g.v3_acc_read(number).clear end
    #8write access write
    puts "8"
    if row_user[0][7] == 'set' then g.v3_acc_write(number).set else g.v3_acc_write(number).clear end
    #9write sources
    puts "9"
    if  row_user[0][8].class == Float
      g.v3_sources(number).set(row_user[0][8].floor.to_s)
    else
      g.v3_sources(number).set(row_user[0][8].to_s)
    end
    #10write inform enable
    puts "10"
    if row_user[0][9] == 'set' then g.v3_notify(number).set else g.v3_notify(number).clear end

    ###############

    if row_user[0][9] == 'set'
      #11write inform heartbeat
      puts "11"
      g.v3_inform_trap(number,row_user[0][10].to_i).set
      #12write destinations
      puts "12"
      if  row_user[0][11].class == Float
        g.v3_destinations(number).set(row_user[0][11].floor.to_s)
      else
        g.v3_destinations(number).set(row_user[0][11].to_s)
      end
      #13write port
      puts "13"
      if  row_user[0][12].class == Float
        g.v3_port(number).set(row_user[0][12].floor.to_s)
      else
        g.v3_port(number).set(row_user[0][12].to_s)
      end
      if row_user[0][10].to_i == 0
        #14write retries
        puts "14"
        if row_user[0][13].class == Float
          g.v3_retries(number).set(row_user[0][13].floor.to_s)
        else
          g.v3_retries(number).set(row_user[0][13].to_s)
        end
        #15write interval
        puts "15"
        if  row_user[0][14].class == Float
          g.v3_interval(number).set(row_user[0][14].floor.to_s)
        else
          g.v3_interval(number).set(row_user[0][14].to_s)
        end
      end
      #16write security level
      puts "16"
      if row_user[0][15] == 'set' then g.v3_heartbeat(number).set else g.v3_heartbeat(number).clear end
    end


    g.save.click

    ######### #read SNMPV3 setting and write result to ss

    sleep 3
    Watir::Waiter.wait_until(5) { g.edit.exists?}
    g.edit.click

    #1read enable user or not
    puts "1"
    if g.v3_enable(number).checked? == true then ws.Range("bc#{row}")['Value']='set'  else ws.Range("bc#{row}")['Value']='clear' end
    #    #2read user name
    puts "2"
    ws.Range("bd#{row}")['Value'] = g.v3_user(number).value.to_s
    #3read authentication
    puts "3"
    ws.Range("be#{row}")['Value'] = g.radio_check(g.det.table(:index, (number*4+1))[2][5]).to_i
    #4read authentication secret
    puts "4"
    ws.Range("bf#{row}")['Value'] = g.v3_auth_secret(number).value.to_s
    #5read privacy
    puts "5"
    ws.Range("bg#{row}")['Value'] = g.radio_check(g.det.table(:index, (number*4+1))[2][6]).to_i
    #6read Privacy secret
    puts "6"
    ws.Range("bh#{row}")['Value'] = g.v3_privacy_secret(number).value.to_s
    #7read access read
    puts "7"
    if g.v3_acc_read(number).checked? == true then ws.Range("bi#{row}")['Value']='set'  else ws.Range("bi#{row}")['Value']='clear' end
    #8read access read
    puts "8"
    if g.v3_acc_write(number).checked? == true then ws.Range("bj#{row}")['Value']='set'  else ws.Range("bj#{row}")['Value']='clear' end
    #9read sources
    puts "9"
    ws.Range("bk#{row}")['Value'] = g.v3_sources(number).value.to_s
    #10read notify 
    puts "10"
    if g.v3_notify(number).checked? == true then ws.Range("bl#{row}")['Value']='set'  else ws.Range("bl#{row}")['Value']='clear' end

    if g.v3_notify(number).checked? == true
      #11read inform heartbeat
      puts "11"
      if   g.v3_inform_trap(number,0).checked? == true then  ws.Range("bm#{row}")['Value'] = 0 else ws.Range("bm#{row}")['Value'] = 1 end
      
      #12read destinations
      puts "12"
      ws.Range("bn#{row}")['Value'] = g.v3_destinations(number).value.to_s
      #13read port
      puts "13"
      ws.Range("bo#{row}")['Value'] = g.v3_port(number).value.to_i
      #14read retries
      puts "14"
      ws.Range("bp#{row}")['Value'] = g.v3_retries(number).value.to_i
      #15read interval
      puts "15"
      ws.Range("bq#{row}")['Value'] = g.v3_interval(number).value.to_i
      #16read heartbeat
      puts "16"
      if g.v3_heartbeat(number).checked? == true  then ws.Range("br#{row}")['Value'] = 'set' else ws.Range("br#{row}")['Value'] = 'clear' end
 
    end

    number +=1
  end

  #  ##Clear all rows of any contents insnmp v3 table
  #  g.edit.click
  #  for i  in 1..20
  #    #puts i
  #    g.v3_clear(i).click
  #  end
  g.save.click
  wb.Save

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
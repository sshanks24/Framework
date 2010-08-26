=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Module_Name*
  Setup

*Description*
  Test script setup methods

*Variables*
    s = test start time
    f = test finish time
    roe = row number in controller spreadsheet if called from controller
    excel = nested array containing an instance of excel and script parameters
      excel[0] = excel instance
        ss = spreadsheet (excel[0][0])
        wb = workbook (excel[0][1])
        ws = worksheet (excel[0][2])
      excel[1] = parameters
        dvr_ss = driver spreadsheet (excel[1][0])
        rows = number of rows in spreadsheet to execute (excel[1][1])
        site = url/ip address of card being tested (excel[1][2])
        name = user name for login (excel[1][3])
        pswd = password for login (excel[1][4])
    ARGV[1] = command line parameter passed from the controller for 'roe'
    (ARGV.length !=0) script called from controller else running independently

=end

module  Setup

  #    - create time stamped controller spreadsheet
  #    - open IE or attach to existing IE session
  #    - populate the spreadsheet with web support page info
  def setup(file)
    base_xl = (file).gsub('/','\\').chomp('rb')<<'xls'
    if(ARGV.length != 0)          # called from controller
      excel = xls_timestamp(base_xl) # called ,connect to existing excel instance
      attach_ie(excel[1][2])  # test site ip
    else
      excel = xls_timestamp(base_xl,'ind') # independent, start new excel instance
      open_ie(excel[1][2])
      support(excel[0])
    end
    return excel
  end


  #    - create time stamped spreadsheet using base name
  #    - connect to excel and open base workbook
  #    - create instance of excel (xl)
  #    - returns a nested array containing spreadsheet and script parameters
  def xls_timestamp(s_s,*type)
    new_ss = (s_s.chomp(".xls")<<'_'<<Time.now.to_a.reverse[5..9].to_s<<(".xls"))
    # common statement assigned with one variable
    if new_ss.include? "controller"
      new_ss1 = new_ss.sub('controller','result')
      xl = new_xls(s_s,1) #open base controller ss with new excel session
    elsif new_ss.include? "driver"
      new_ss1 = new_ss.sub('driver','result')
      if (type.to_s == 'ind') # driver was launched independently
        xl = new_xls(s_s,1) #open base driver ss with new excel session
      else # driver was launched by controller
        xl = conn_open_xls(s_s,1) #connect to existing excel session and create new workbook for L2
      end
    end

    ws = xl[2] # worksheet

    param = Array.new
    param[0] = new_ss1
    param[1] = ws.Range("b2")['Value'].to_i        # rows
    param[2] = ws.Range("b3")['Value']             # Test_site
    param[3] = ws.Range("b4")['Value']             # username
    param[4] = ws.Range("b5")['Value']             # password

    # This is a nested array that contains the instance of excel
    # and the spreadsheet parameters listed directly above
    excel = [xl,param]

    # save spreadsheet as timestamped name.
    save_as_xls(xl,new_ss1)
    return excel
  end


  #
  #    - create new IE instance and navigate to the test site
  def open_ie(site)
    puts "\n    **Open IE **\n"
    $ie = Watir::IE.new
    $ie.goto(site)
   end


  #
  #     - attach to existing IE instance and navigate to the test site
  def attach_ie(site)
    puts "\n    **Attach to IE **\n"
    site = 'http://'<<site<<'/'
    $ie = Watir::IE.attach(:url, site)
  end


  #
  #    - navigate with 'click'; script is called from controller; no login req'd
  #    - navigate with 'click_no_wait'; script called directly; login req'd
  def logn_chk(nav,logn)
    site,name,pswd = logn[2,4] #logn is an array
    login(site,name,pswd) if (ARGV.length == 0)
    nav.click
  end

  
  #    - thread based method to login to web page
  #    - acknowledge security alert
  #    - uses standard .click as opposed to .click_no_wait
  def login(site,user,pswd)
    conn_to = 'Connect to '+ site
    Thread.new{
      thread_cnt = Thread.list.size
      sleep 1 #This sleep is critical, timing may need to be adjusted
      Watir.autoit.WinWait(conn_to)
      Watir.autoit.WinActivate(conn_to)
      Watir.autoit.Send(user)
      Watir.autoit.Send('{TAB}')
      Watir.autoit.Send(pswd)
      popup('Windows Internet Explorer','OK') #launch thread for alert popup
      Watir.autoit.Send('{ENTER}')
    }
  end

  def popup(name,btn) #alert popup thread
    Thread.new{
      sleep 1 # This sleep is critical, timing may need to be adjusted
      Watir.autoit.WinWait(name)
      Watir.autoit.ControlClick(name,"",btn)
    }
  end

   
  #    - collects support page table:
 
  def support(xl)
    puts "  Collect Support page info"
    supp.click
    sleep 1
    wb,ws = xl[1,2]
    row = 11
    supprt.each do|key|
      if !key[0].nil?
        c = ws.range("A#{(row)}:B#{(row)}")
        c.value = key
        c.Interior['ColorIndex'] = 19   # change background color
        c.Borders.ColorIndex = 1        # add border
        #ws.range("A#{row}:B#{row}").Columns.Autofit
        row+=1
      end
      ws.range("A:B").Columns.Autofit
    end
    wb.Save
  end

  def snmp_setup(wb)
    ws = wb.Worksheets('SupportInfo')
    @test_site = ws.range('B3').value
    @community_string = ws.range('B5').value # Use the password field for the community string
  end
           
end

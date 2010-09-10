
=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Class_Name*
  Generic_Parent  this is inherited by Generic

*Description*
  common methods used by most scripts

*Variables*
  start     = first table row
  row       = last table row
  col       = table column
  idx       =
  ws        =
  href_rst  =
  a         =
  button    =
  user_inp  =
  ttxt      =
  pop_exp   =
  pop_name  =
  


    
=end

$:.unshift File.dirname(__FILE__) unless
$:.include? File.dirname(__FILE__)

require 'watir/ie'
require 'watir\contrib\enabled_popup' #This is called by jsClick
require 'time'
require 'win32ole'
require 'navigate12'
require 'xls4'
require 'setup29'
require 'teardown14'
require 'snmp1'
require 'telnet16'
require 'bacnet'


class Generic_parent
  include Nav
  include Xls
  include Teardown
  include Setup
  include Snmp
  include Telnet_cstm

 
  #
  # - used for configure information table scripts
  # - 'fw == 1' is for firmwareweb_table_info
  def table_info(start,_end,col,idx,ws,fw=nil)
    # iterate through all rows and columns
    while (start <= _end)
      j = 1
      while (j <= col)
        case j
        when 1
          if fw == 1
            parameter = param_firmwrweb(idx,start,j).text
          else
            parameter = param_descr(idx,start,j).text
          end
          ws.Range("bc#{start+1}")['Value'] =  parameter
        when 2
          if fw == 1
            description = param_firmwrweb(idx,start,j).text
          else
            description = param_descr(idx,start,j).text
          end
          ws.Range("bd#{start+1}")['Value'] = description
        end
        j += 1
      end
      start += 1
    end
  end


  #
  # - currently not used, will require refactoring for latest framework
  def restart(gen,href_rst) #TODO: this method requires refactoring
    puts "click on Restart folder in the left pane"
    #sleep 11
    puts"href_restart = #{href_rst}"
    #link(gen,'imgConfigure','Restart')
    #link_name='Restart' or 'Reinitialize'

    restart_link = $ie.frame(:index, 3).frame(:index, 2).link(:href,href_rst).exists?
    puts "restart_link = #{restart_link}"
    if (restart_link == false)
      sleep 7
      main_tab('imgConfigure')
      restart_link = $ie.frame(:index, 3).frame(:index, 2).link(:href,href_rst).exists?
      puts "restart_link = #{restart_link}"
    end
    $ie.frame(:index, 3).frame(:index, 2).link(:href,href_rst).click_no_wait
    login_pop_exists = login($test_site,$login,$password)
    if (login_pop_exists == '1')
      jsClick($ie,"OK")
    end

    puts "Click on the Restart Button in the Restart dialogue"

    # Click_no_wait is needed at this point or the script will loose control
    # when the card restarts.
    $ie.frame(:index, 3).frame(:index, 3).button(:name,'Submit').click_no_wait
    puts "go to sleep for 60 seconds"
    sleep 30 #TODO replace the hard sleeps with ping polling method
    puts "I'm awake now. Let's go back to the test-site!"

    $ie.close
    sleep 40
  end

    
  #  - function jsClick- Handle popup and return pop up text if 'rtxt' is true
  #  - user_input is used for firmware update file dialogue box
  def jsClick( a, button, user_inp = nil,rtxt = nil)
    wait = 70
    hwnd1 = $ie.enabled_popup(wait) # wait up to 60 seconds for a popup to appear
    #puts "hwnd = #{hwnd1}"
    if (hwnd1)
      w = WinClicker.new
      if (rtxt)
      popup_text = w.getStaticText_hWnd(hwnd1).to_s.delete "\n"
      end
      if (user_inp)
        w.setTextValueForFileNameField(hwnd1, "#{user_inp}")
      end
      sleep (0.1)
      w.clickWindowsButton_hwnd(hwnd1, "#{button}")
      w = nil
    end
    return popup_text
    puts"pop-up text = #{popup_text}"
  end

  
  #  - after attempting to save an invalid character - reset OK or
  #  - reset Cancel And reset OK, return text in popup
  def invChar( a,pop_exp,user_inp = nil)
    save.click_no_wait
    poptxt = jsClick( $ie,"OK",user_inp = nil,"rtxt")
    if (pop_exp == "can")
      reset.click_no_wait
      jsClick( $ie,"Cancel",user_inp = nil)
      edit.click
    end
    reset.click_no_wait
    jsClick( $ie,"OK",user_inp = nil)
    return poptxt
  end  
  
  
  #  
  #  - reset Cancel or reset OK, implicitly return text in popup
  def res_can(pop_exp)
    if (pop_exp == "res")
      reset.click_no_wait
      jsClick( $ie,"OK",user_inp = nil,"rtxt")
    elsif (pop_exp == "can")
      reset.click_no_wait
      jsClick( $ie,"Cancel",user_inp = nil,"rtxt")
    end
  end  


  #  - reset to factory defaults ok or cancle, return text in popup
  def res_factory(pop_exp)
    if (pop_exp == "res")
      restart1.click_no_wait
      jsClick( $ie,"OK",user_inp = nil,"rtxt")
    elsif (pop_exp == "can")
      restart1.click_no_wait
      jsClick( $ie,"Cancel",user_inp = nil,"rtxt")
    end
  end
  
  #    
  #  - read checkbox status and return set of clear
  def checkbox(box)
    if box.checked?
      'set'
    else
      'clear'
    end
  end

    #-read radio status and return radio number which selected in the table
  def radio_check(table)
    table.radios.each{|x| if x.checked? then  return  x.value end }
  end

    #  - <code>frame_idx</code> is the numerical index of the frame you want to
  #  - check for tables.
  def has_table?(frame_idx)
    str = ''
    $ie.frame(:index,frame_idx).tables.each do |t|
       str << t.to_s
    end
    return !str.empty?
  end

  #  - returns true or false if the web page under test has a frame named
  #  - <code>frame_name</code>
  def has_frame?(frame_name)
    frame_text = self.redirect {$ie.show_frames}
    !frame_text.match("#{frame_name}").nil?
  end

  #  - writes table contents in frame <code>frame_idx</code> (less header row)
  #  - to worksheet <code>ws</code>
  def table_to_ss(frame_idx, ws, report_name)
    if self.has_table?(frame_idx)
      $ie.frame(:index,frame_idx).tables.each do |table|
        if table.row_count > 1 #Header tables only have 1 row - we don't want those.
          for i in 2..table.row_count
            ws.cells(@row_ptr, 1).value = report_name.to_s #Writes the link(report)
            for j in 2..table.column_count
              if report_name =~ /\[.*\]/ and j == 2 then # Special case for multi-modules
               ws.cells(@row_ptr, j).value = table[i][j].text + " (#{report_name})"
              else
                ws.cells(@row_ptr, j).value = table[i][j].text
              end
            end
            @row_ptr += 1
          end
        end
      end
    end
  end

  #  - returns a compacted array (no nil values)
  #  - Variables
  #  - title - The textual representation of the title, will match a regex
  #  - frame_idx - numeric index of the frame in which the table of interest
  #  - belongs.
  #  - title_pos - the position in the table in which the 'title' lies.
  #  - Note that tables indices start at 1.  Ruby arrays index start at 0...
  def get_table_by_title(title,frame_idx=3,title_pos=[1,2])
    $ie.frame(:index,frame_idx).tables.each do |table|
      if table[title_pos[0]][title_pos[1]].text =~ /#{title}/
        return table.to_a.compact
      end
    end
    return nil
  end

  # Stops execution until 20 seconds after a ping echo reply is received
  # Useful
  def wait_for_reboot(ip)
    puts "\n\n"
    flag = true
    reply_from = Regexp.new('Reply from')
    while flag == true
      puts "Waiting for device to reboot"
      sleep(20)
      results = `ping #{ip}`
      puts results
      if (reply_from.match(results)) then
        puts "Device is booting..."
        sleep(20)
        flag = false
      end
    end
    puts "\n\n"

  end

end

class Array
  #  - An attempt to generalize outputting an array of values to a result spread
  #  - sheet.
  #  - Variables
  #  - ws - the worksheet to write the array to
  #  - depth - for multi-dimensional arrays - think # of columns
  #  - row_start -
  #  - col_start -
  def to_spread_sheet(ws,depth=1,row_start=1,col_start=1)
    tmp = self.flatten
    tmp.each_slice(depth) do |row|
      col = col_start
      row.each do |item|
        ws.cells(row_start, col).value = item.to_s
        col +=1
      end
      row_start += 1
    end
  end
end

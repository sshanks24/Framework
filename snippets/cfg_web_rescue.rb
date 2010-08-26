=begin
#-------------------------------------------------------------------#
18-Oct-2009
This snippet is designed to demonstrate and test recovery mechanisms 
for the web card. On occassion the script seems to get lost at various 
points. Various rescue statements are used to restore script execution. 

This script is in sync with: IS-WEBCARD_HID6_3.100.2_048951   
#-------------------------------------------------------------------#
=end



$:.unshift File.dirname(__FILE__) unless
$:.include? File.dirname(__FILE__)

require 'watir/ie'
require 'win32ole'
require 'watir\contrib\enabled_popup'

def modal(pop_name,button) # modal alert popup thread
  Thread.new{
    sleep 1 # This sleep is critical, timing may need to be adjusted
    Watir.autoit.WinWait(pop_name)
    Watir.autoit.ControlClick(pop_name,"",button)
    puts " Modal Alert popup completed"
  }
end


def fastLogin(test_site,user,pswd) # thread based method
  conn_to = 'Connect to '+ test_site
  Thread.new{
    puts" #{conn_to}, \n #{user}, \n #{pswd}"
    sleep 1 #This sleep is critical, may need to be adjusted
    Watir.autoit.WinWait(conn_to)
    Watir.autoit.WinActivate(conn_to)
    Watir.autoit.Send(user)
    Watir.autoit.Send('{TAB}')
    Watir.autoit.Send(pswd)
    modal('Windows Internet Explorer','OK') #launch thread to handle modal popup
    Watir.autoit.Send('{ENTER}')
    puts " Login completed"
  }
end


def jsClick( a, button, user_input=nil)# java script popup method
 waitTime = 30
 begin
   hwnd = $ie.enabled_popup(waitTime)
    if (hwnd)
      w = WinClicker.new
      if ( user_input )
        w.setTextValueForFileNameField(hwnd, "#{user_input}")
      end
      sleep 3
      w.clickWindowsButton_hwnd(hwnd, "#{button}")
      w = nil
    end
  rescue Timeout::Error
  puts "there was no popup within the wait time"
  end
end


# Locally define only the objects that will be used in this script

# define the tabs
def config
  $ie.frame(:index, 2).image(:id, 'imgConfigure')
end


# define links  
def cfgweb 
  $ie.frame(:index, 3).frame(:index, 2).link(:text => 'Web', :index => 2)
end


# define edit,save,reset buttons
def edit
  $ie.frame(:index, 3).frame(:index, 3).button(:id, 'editButton')
end

def save
  $ie.frame(:index, 3).frame(:index, 3).button(:name, 'Submit')
end

def reset 
  $ie.frame(:index, 3).frame(:index, 3).button(:value, 'Reset')
end


# define cfg_web Controls
def cfgCtrl
  $ie.frame(:index, 3).frame(:index, 3).checkbox(:id, 'enableWebCfgCtrl')
end

test_site = "10.130.92.34"
user = 'Liebert'
pswd = 'Liebert'
conn_to = 'Connect to '+ test_site

# $HIDE_IE=true 
$ie = Watir::IE.new
$ie.goto test_site
$ie.maximize

puts " - Start watir/webcard endurance test"

start = Time.now
config.click
fastLogin(test_site,user,pswd) #Start login thread
cfgweb.click # click_no_wait not needed when login thread is started before hand 

sleep 1


x = 0 # initialize iteration count
pass = 0 # initialize success count
fail = 0 # initialize error count

while x < 5
  x +=1
  puts "Test step =#{x}"
  begin 
    puts "Start of normal loop"
    s = Time.now
    sleep 2
    Watir::Waiter.wait_until(10) {edit.exists?}
    edit.click
    Watir::Waiter.wait_until(10) {save.exists?}
    save.click_no_wait
    sleep 1  # need to remove this sleep to see what the effect is
    jsClick( $ie, "OK")
    f = Time.now
    e = (f-s)
    pass +=1
    puts "Success - #{pass} of #{x} @ #{e} seconds"
   rescue
    fail +=1
    puts "Error - #{fail} of #{x}"
    Watir.autoit.Send('{F5}')
    puts"********** rescue ***"
    sleep 2
    begin
      sleep 5 # wait for the previous click_no_wait process to time out before starting another 
      save.click_no_wait
      jsClick( $ie, "OK")
    rescue
      puts "**********no 'save' popup in rescue***"
    end
    config.click
    begin
      cfgweb.click_no_wait
      jsClick( $ie, "OK")
      puts"**********cfgWeb popup ***"
    rescue
      puts"**********return-1 to normal loop***"
    end
    puts"**********return-2 to normal loop***"
    x = x-1
  end
end

done = Time.now
total = (done-start)
puts "Total time = #{total}"
$ie.close


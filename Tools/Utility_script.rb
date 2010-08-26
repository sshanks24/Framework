=begin
#--------------------------------------------------------------------------------#
This is Controller script which will be launched first.

 # xL_param is a nested array that contains two arrays
 # xL_param[0] = Time stamped spreadsheet
 # xL_param[1] = An array containing spreadsheet object and data is returned

 #--------------------------------------------------------------------------------#
=end

lib = File.dirname(__FILE__).sub('Tools','lib') # add lib to load path
dvr = File.dirname(__FILE__).gsub('controller','driver') 
$:.unshift lib
$:.unshift dvr
s = Time.now 
require 'generic'

def log_cred
    puts "Enter the IP address "
    ip = gets
    ip = ip.strip
    puts "The device on which the script is to be run is #{ip}"
    puts "Enter the username"
    u_name = gets
    u_name = u_name.strip
    puts "Enter the password"
    pas_wd = gets
    pas_wd = pas_wd.strip
    puts "Enter the Controller spreadsheet number: example - 46:"
    ss1_no = gets.chomp!
    cred = Array.new
    cred = [ip,u_name,pas_wd,ss1_no]
    return cred
  end

begin
  puts" \n Executing: #{(__FILE__)}\n\n" #current file
  g = Generic.new
  file = (__FILE__.sub('Tools','Controller')).sub('Utility_script3.rb','Controller_46.rb')
 
  cntrlr_ssname = file.sub('.rb','.xls')
  puts cntrlr_ssname
  xls1 = g.new_xls(cntrlr_ssname,1)
  ss1  = xls1[0]
  wb1  = xls1[1]
  ws1  = xls1[2]
  rows = ws1.Range("b2")['Value']
  puts"rows = #{rows}"
  cred = log_cred
  ip     = cred[0]
  login  = cred[1]
  passwd = cred[2]
  row = 1 
    
  while (row <= rows)
    r_ow = row # script number
    row += 1 # start at row 2
    run_flag = ws1.Range("e#{row}")['Value']
    puts "run_flag = #{run_flag}"
    if run_flag == true
      print" Executing Driver script #{r_ow} -- "
      file_dir = File.dirname(file)
      path = file_dir.sub('Controller','driver/') # driver path
      drvr = path << (ws1.Range("j#{row}")['Value'].to_s) #concat path and driver 
      ss_name = drvr.sub('.rb','.xls')
      puts ss_name
      
      xls2 = g.new_xls(ss_name,1)
      
      ss2  = xls2[0]
      wb2  = xls2[1]
      ws2  = xls2[2]
      
      ws2.Range("b3")['Value'] = ip
      ws2.Range("b4")['Value'] = login
      ws2.Range("b5")['Value'] = passwd
      
      wb2.Save
      ss2.Quit
      g.conn_act_xls # connect to controller spreadsheet 
    end
  end
    ws1.Range("b3")['Value'] = ip
    ws1.Range("b4")['Value'] = login
    ws1.Range("b5")['Value'] = passwd
    wb1.Save
    
  wb1.Save
  ss1.Quit
  f = Time.now
  
  
rescue Exception => e
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
ensure 
  
end


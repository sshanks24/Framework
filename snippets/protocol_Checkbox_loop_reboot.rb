$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
require 'lib\generic'
require 'watir'


def wait_for_reboot(ip)
  puts "\n\n"
  flag = true
  reply_from = Regexp.new('Reply from')
  while flag == true
    puts "Waiting for device to reboot"
    sleep(10)
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

begin
  loop_time = 250
  ip = "126.4.202.203"
  user = [' ',' ',"#{ip}",'Liebert','Liebert'] #because Logn_chk start from Array[2]

  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  g.open_ie(ip)
  $ie.speed = :zippy
 m1,m2,m3 = 0,0,0  #define 3 flag to record if checkbox is exist
  #####loop start
  for i in 1..loop_time
    g.config.click
    g.logn_chk(g.mgtprot,user)
    if g.edit.exist?
      g.edit.click
      if  g.checkbox( g.snmp_agent)  then m1 += 1 end
      if  g.checkbox(g.snmp_v1v2)    then m2 += 1 end
      if  g.checkbox(g.snmp_v3)        then m3 += 1 end
    else
      puts "Management protocol edit button is not exist!"
    end
    puts " loop time = #{i}, snmp chickbox loop time = #{m1}, #{m2},#{m3}"

    #reset to defalt factury
    g.factdef.click
    g.res_factory("res" )
    
    #wait until restart finished
    s = Time.now
    wait_for_reboot(ip)
    f = Time.now

    puts (f-s).to_s
    $ie. goto ip
  end
end
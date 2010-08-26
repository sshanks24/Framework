$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
require 'lib\generic'
require 'watir'


def wait_for_reboot(ip)
  #  puts "\n\n"
  flag = true
  reply_from = Regexp.new('Reply from')
  while flag == true
    puts "Waiting for device to reboot"
    sleep(10)
    results = `ping #{ip}`
    #    puts results
    if (reply_from.match(results)) then
      puts "Device is booting..."
      sleep(20)
      flag = false
    end
  end
  puts "\n"
end

begin
  loop_time = 250
  ip = "126.4.202.207"
  user = [' ',' ',"#{ip}",'Liebert','Liebert'] #because Logn_chk start from Array[2]

  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  g.open_ie(ip)
  $ie.speed = :zippy
  m1,m2,m3 = 0,0,0  #define 3 flag to record if checkbox is exist
  #####loop start
  for i in 1..loop_time
    puts "loop start"

    # 1.check management protocol checkboxes
    puts "1"
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


    # 2.restore to factory defaults
    puts "2"
    g.factdef.click
    g.res_factory("res" )


    # 3.wait until restart finished
    puts "3"
    s = Time.now
    wait_for_reboot(ip)
    f = Time.now
    puts (f-s).to_s
    #$ie. goto ip
    p 'goto google'
    $ie. goto 'www.google.com'
    p 'goto ip'
    $ie. goto ip


    # 4.creat snmp3 user
    puts "4"
    sleep 4
    g.config.click
    g.logn_chk(g.snmpv3,user)
    g.edit.click
    g.v3_clear(1).click
    number = 1
    #################write snmp user info############
    #1write enable user or not
    g.v3_enable(number).set
    #2write user name
    g.v3_user(number).set("user1")
    #3write authentication
    g.v3_auth(number, 0).set
    #4write authentication secret
    g.v3_auth_secret(number).set("testtest123456")
    #5write privacy
    g.v3_privacy(number,0).set
    #6write Privacy secret
    g.v3_privacy_secret(number).set("test123456")
    #7write access read
    g.v3_acc_read(number).set
    #8write access write
    g.v3_acc_write(number).set
    #9write sources
    g.v3_sources(number).set("126.4.100.100")
    #10write inform enable
    g.v3_inform_enable(number).set
    #11write inform heartbeat
    g.v3_inform_heartbeat(number).set
    #12write destinations
    g.v3_destinations(number).set("126.4.100.100")
    #13write port
    g.v3_port(number).set("165")
    #14write retries
    g.v3_retries(number).set("3")
    #15write interval
    g.v3_interval(number).set("15")
    #16write security level
    g.v3_security_level(number,0).set

    g.save.click
    sleep 5
    Watir::Waiter.wait_until(30) { g.edit.exists?}

    # 5.disable v1/v2 protocal
    puts "5"
    g.mgtprot.click
    g.edit.click
    g.snmp_v1v2.clear
    g.save.click_no_wait

    # 6.restore to factor defaults
    puts "6"
    g.factdef.click
    g.res_factory("res" )

    # 7.wait until restart finished
    puts "7"
    s = Time.now
    wait_for_reboot(ip)
    f = Time.now
    puts (f-s).to_s
    p 'goto google'
    $ie. goto 'www.google.com'
    p 'goto ip'
    $ie. goto ip
    puts "loop end"
  end
end
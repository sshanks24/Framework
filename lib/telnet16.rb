=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Module_Name*
  Telnet_cstm

*Description*
  telnet methods

*Variables*
  name = user name for login
  pswd = password for login

telnet1-close function changed with parameters

telnet 2-modified telnet_op_read with exit
3- testing for write
6-One common function whcih uses all gven functions
7-One more common function added to have read and write separately
13-op write and read has main back as part of function
=end

require 'net/telnet'
require 'watir/ie'  #TODO watir should not be required here

#TODO documentation of methods and variables needs to be completed
#TODO implement newer methods from the keyword driven model
module Telnet_cstm

  def telnet_connect(ip,login,pswd)
    prompt_type=/[ftpusr:~>]*\z/n
    puts"ip,=#{ip},login=#{login},password=#{pswd}"
    telnet = Net::Telnet.new('Host' => ip,'Prompt' =>prompt_type ,"Output_log"   => "dump_log.txt"  )

    #The prompt should be the real prompt while you entered your system
    telnet.cmd(''){|c| print c}
    #telnet.cmd('df -k'){|c| print c}
    telnet.waitfor(/[Ll]ogin[: ]*\z/n) {|c| print c}
    telnet.cmd(login) {|c| print c}
    telnet.waitfor(/Password[: ]*\z/n) {|c| print c}
    telnet.cmd(pswd) {|c| print c}

    telnet.waitfor(/[>]/n) {|c| print c}
    #stdout = telnet.read_until(/[ftpusr:~>]*\z/n)
    sleep 5
    return telnet
  end

  def telnet_op_read(telnet,path,gen,row,*level)
    #level=level.to_s
    spreadsheet=gen.open_xls(path,1)
    worksheet=spreadsheet[2]

    level_count=level.length-1
    #puts"level_count=#{level_count}"
    iteration=0
    output=Array.new

    while (iteration <=level_count)
      #puts"telnet connected=#{p}"
      if(iteration == level_count)
        worksheet.Range("bt#{row}")['Value']=output
      end
      telnet.cmd(level[iteration]){|c| print c}
      sleep 5
      output[iteration]=telnet.waitfor(/[>]/n) {|c| print c}
      iteration+=1
    end

    gen.close_xls(spreadsheet)
    level_count+=1
    back_main(telnet,level_count)
    #return level_count+1

  end

  #iteration[0] value indicates if op write will be made more than once indicated by 1
  def telnet_op_write(telnet,*level)
    level_count=level.length
    iteration=1
    while (iteration <(level_count-1))
      telnet.print(level[iteration]){|c| print c}
      sleep 5
      iteration+=1
      if (iteration ==(level_count-1))
        #level_count=3
        #puts "iteration=#{iteration},value= #{level[iteration]},level_count=#{level_count}"
        telnet.cmd(level[iteration]){|c| print c}
      else
        telnet.waitfor(/[>]/n) {|c| print c}
      end
    end

    if (iteration == 'yes')
      level_count =level_count-2
    else
      level_count =level_count-3
    end
    #puts"level_count at write exit =#{level_count}"
    back_main(telnet,level_count)
    #return level_count
  end

  def back_main(telnet,level_count)
    crsr_up = "\033"
    iteration=0
    while (iteration <(level_count))
      telnet.cmd(crsr_up){|c| print c}
      sleep 5
      iteration+=1
      if (iteration !=(level_count))
        telnet.waitfor(/[>]/n) {|c| print c}
      end
    end
  end

  def telnet_save_exit(telnet)

    #telnet.waitfor(/[>]/n) {|c| print c}
    telnet.print('x'){|c| print c}
    sleep 30
    telnet.waitfor(/[>]/n) {|c| print c}
    telnet.close
    print"\ndone"
  end

  #,number_count- defines number of times u want to do write operation in same telnet session
  def telnet_custom_write(gen,*level)
    telnet=gen.telnet_connect($test_site,$login,$password)
    gen.telnet_op_write(telnet,*level)
    telnet_save_exit(telnet)
    sleep 30
  end

  def telnet_custom_read(gen,path,row,*level)
    telnet=gen.telnet_connect($test_site,$login,$password)
    level_count=gen.telnet_op_read(telnet,path,gen,row,*level)
    telnet_save_exit(telnet)
  end

end

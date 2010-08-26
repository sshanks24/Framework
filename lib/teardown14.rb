=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Module_Name*
  Teardown

*Description*
  Test script teardown methods

*Variables*
    s = test start time
    f = test finish time
    roe = row number in spreadsheet 
    row = used to determine if driver or controller spreadsheet
    s_s = excel instance
      ss = spreadsheet
      wb = workbook
      ws = worksheet
=end


module  Teardown

  #
  #  - teardown driver - this function will update driver spreadsheet.
  def tear_down_d(s_s,s,f,roe)
    # The variable 's_s' is an array that holds the current spreadsheet instance 
    ss,wb,ws = s_s
    #Save Summary and elapsed time to current ss
    ws.Range("b6")['Value'] = s.to_s
    ws.Range("b7")['Value'] = f.to_s
    run_time = elapsed(f,s)
    ws.Range("b8")['Value'] = run_time.to_s
    status = ws.Range("b9")['Value'].to_s # Pass / Fail from Driver.xls
    puts "Status      = #{status}"
    wb.save
    wb.close #Close Driver spreadsheet
    if roe > 0  # called from controller
      # Connect to the controller script to update the status and time
      ss = WIN32OLE.connect('excel.Application')
      wb = ss.ActiveWorkbook
      ws = wb.Worksheets(1)
      ws.Range("l#{roe}")['Value'] = status.to_s
      ws.Range("n#{roe}")['Value'] = run_time.to_s
      wb.save
    end  
  end

  
  #
  #  - teardown controller - this function will update controller spreadsheet.
  def tear_down_c(xl,s,f)
    # The variable 's_s' is an array that holds the current spreadsheet instance 
    ss,wb,ws = xl
    #Save Summary and elapsed time to current ss
    ws.Range("b6")['Value'] = s.to_s
    ws.Range("b7")['Value'] = f.to_s
    run_time = elapsed(f,s)
    ws.Range("b8")['Value'] = run_time.to_s
    wb.save
    wb.close #Close Driver spreadsheet
    ss.quit
    $ie.close
  end
  
  
  #
  #  - calculates difference between start and finish time
  def elapsed(finish,start,*row)
    time = (finish-start).to_i
    hours = time/3600.to_i
    minutes = (time/60 - hours * 60).to_i
    seconds = (time - (hours * 3600 + minutes * 60)).to_i
    if(row != 0 ) # If driver script - use min and sec
      test_time  = minutes.to_s << 'min'<<seconds.to_s<<'sec'
      puts "\n\nTest Start  = "<<start.strftime("%H:%M:%S")
      puts "Test Finish = "<<finish.strftime("%H:%M:%S")
      puts "Test Time   = #{minutes}min#{seconds}sec"
    else # Default output (Controller script) use hrs, min, and sec
      test_time  = hours.to_s << 'hr' <<minutes.to_s << 'min'<<seconds.to_s<<'sec'
      puts "\n\nTest Start  = "<<start.strftime("%H:%M:%S")
      puts "Test Finish = "<<finish.strftime("%H:%M:%S")
      puts "Test Time   = #{hours}hr#{minutes}min#{seconds}sec"
    end  
    return test_time
  end

  
  #
  #  - can be called to kill excel.exe
  def kill()
    kill_result = %x{tskill excel}
    #kill_result = system"TASKKILL /F /IM EXCEL.EXE"
  end  
end  
  
  
  
  


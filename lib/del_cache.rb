#!/bin/ruby
#
#
#
#
#
# set TEMPIF=%USERPROFILE%\Local Settings\Temporary Internet Files
# %TEMPIF%\Content.IE5\Index.DAT

#TODO This file needs to be cleaned up
require 'Win32API'

#
#
# HashMethods = a Hash that can be accessed with methods
#   e.g.    h = MethodHash.new
#           h['street'] = 'Broadway'
#           h.street    = 'Broadway'
#           puts h.street  ===> Broadway

require 'delegate'

class MethodHash < SimpleDelegator
    def initialize h = {}
        super h
    end
    
    def method_missing(method_name, *args)
        name = method_name.to_s
        if name.ends_with?('=')
            self[ name.chop ] = args[0]
        else
            self[ name ]
        end
    end
end


class String
    def ends_with?(substr)
        len = substr.length
        self.reverse() [0 .. len-1].reverse == substr
    end
    
    def starts_with?(substr)
        len = substr.length
        self[0 .. len-1] == substr
    end
    
    alias start_with?  starts_with?
    alias begin_with?  starts_with?
    alias begins_with? starts_with?
    alias end_with?    ends_with?

    # String each() operator reads line-by-line
    # These functions return characters
    def each_char
        self.each_byte{|x| yield x.chr }
    end
    def collect_char
        r = []
        self.each_byte{|x| r << x.chr }
        r
    end
end


=begin

// Windows System calls needed to clear cache.
//

BOOL DeleteUrlCacheEntry(
  LPCTSTR lpszUrlName
);

HANDLE FindFirstUrlCacheEntry(
  LPCTSTR lpszUrlSearchPattern,
  LPINTERNET_CACHE_ENTRY_INFO lpFirstCacheEntryInfo,
  LPDWORD lpdwFirstCacheEntryInfoBufferSize
);

BOOL FindNextUrlCacheEntry(
  HANDLE hEnumHandle,
  LPINTERNET_CACHE_ENTRY_INFO lpNextCacheEntryInfo,
  LPWORD lpdwNextCacheEntryInfoBufferSize
);

BOOL FindCloseUrlCache(
    IN HANDLE hEnumHandle
);


typedef struct _INTERNET_CACHE_ENTRY_INFO {
    DWORD dwStructSize;
    LPTSTR lpszSourceUrlName;
    LPTSTR lpszLocalFileName;
    DWORD CacheEntryType;
    DWORD dwUseCount;
    DWORD dwHitRate;
    DWORD dwSizeLow;
    DWORD dwSizeHigh;
    FILETIME LastModifiedTime;
    FILETIME ExpireTime;
    FILETIME LastAccessTime;
    FILETIME LastSyncTime;
    LPBYTE lpHeaderInfo;
    DWORD dwHeaderInfoSize;
    LPTSTR lpszFileExtension;
    union {
        DWORD dwReserved;
        DWORD dwExemptDelta;
    }
} INTERNET_CACHE_ENTRY_INFO, *LPINTERNET_CACHE_ENTRY_INFO;


=end




def main
    w = get_api('wininet',FUNCS)
    i = 0
    
    info,infosize = get_first_info(w)
    cache = w.FindFirstUrlCacheEntry.Call(nil,info,infosize)
    if cache != 0
        begin
            len, source_file_ptr, local_file_ptr = info.unpack 'LLL'
            w.DeleteUrlCacheEntry.Call(source_file_ptr)
            i += 1
            info,infosize = get_next_info( w, cache )
        end while w.FindNextUrlCacheEntry.Call(cache, info, infosize) != 0
    end
    w.FindCloseUrlCache.Call(cache)
    i
end

def get_first_info(api)
    sizenum = [0,0].pack('L*')
    buf =  [1024,0].pack('L*').ljust(1024)
    r = api.FindFirstUrlCacheEntry.Call(nil,nil,sizenum)
    n = sizenum.unpack('L')[0]
    info = sizenum.ljust(n)
    [info,sizenum]
end

def get_next_info(api, handle)
    sizenum = [0,0].pack('L*')
    buf =  [1024,0].pack('L*').ljust(1024)
    r = api.FindNextUrlCacheEntry.Call(handle,nil,sizenum)
    n = sizenum.unpack('L')[0]
    info = sizenum.ljust(n)
    [info,sizenum]
end


#
# Win32 API used in this file is listed here.
# Each system call can be instatiated like this
#    DeleteUrlCacheEntry = Win32API.new("wininet", "DeleteUrlCacheEntry", ['P'], 'V')
# Instead, the functions are defined dynamically from this list:
#
FUNCS = {
'FindFirstUrlCacheEntry' => ['ppp','n'],
'FindNextUrlCacheEntry'  => ['npp','n'],
'DeleteUrlCacheEntry'    => ['n',  'n'],
'FindCloseUrlCache'      => ['n',  'n']
}

#
# get_api returns a hash with Win32 API system calls
# Usage:
#     api = get_api('Kernel32', {'GetLastError'=>['V', 'N'],...})
#
def get_api(library, function_hash)
    f = MethodHash.new
    function_hash.each{|funcname,types|
        in_types = types[0].collect_char{|x| x}
        out_type = types[1]
        f[funcname] = Win32API.new(library, funcname, in_types, out_type)
    }
    f
end


def getLastError
    f = Win32API.new('Kernel32', 'GetLastError', ['V'], 'N')
    f.call
end


if __FILE__ == $0
    n = main
    # print "Deleted #{n} items\n"
end


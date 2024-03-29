import socket
import struct
import datetime
import win32api
from pyuac import main_requires_admin # pip install pyuac

# List of servers in order of attempt of fetching
server_list = ['time.windows.com']

def gettime_ntp(addr='time.windows.com'):
    '''
    Returns the epoch time fetched from the NTP server passed as argument.
    Returns none if the request is timed out (5 seconds).
    '''
    # http://code.activestate.com/recipes/117211-simple-very-sntp-client/
    TIME1970 = 2208988800      # Thanks to F.Lundh
    client = socket.socket( socket.AF_INET, socket.SOCK_DGRAM )
    data = bytes('\x1b' + 47 * '\0', 'utf-8')
    try:
        # Timing out the connection after 5 seconds, if no response received
        client.settimeout(5.0)
        client.sendto( data, (addr, 123))
        data, address = client.recvfrom( 1024 )
        if data:
            epoch_time = struct.unpack( '!12I', data )[10]
            epoch_time -= TIME1970
            return epoch_time
    except socket.timeout:
        return None

@main_requires_admin
def updateTime(time=datetime.datetime.now()):
    '''
    Updates the time of the system to the time passed as argument.
    This function requires admin privileges.
    '''
    win32api.SetSystemTime(utcTime.year, utcTime.month, utcTime.weekday(), utcTime.day, utcTime.hour, utcTime.minute, utcTime.second, 0)
    # Local time is obtained using fromtimestamp()
    localTime = datetime.datetime.fromtimestamp(epoch_time)
    print("Time updated to: " + localTime.strftime("%Y-%m-%d %H:%M") + " from " + server)

if __name__ == "__main__":
    # Iterates over every server in the list until it finds time from any one.
    for server in server_list:
        epoch_time = gettime_ntp(server)
        if epoch_time is not None:
            # SetSystemTime takes time as argument in UTC time. UTC time is obtained using utcfromtimestamp()
            utcTime = datetime.datetime.utcfromtimestamp(epoch_time)
            print(utcTime)
            updateTime(utcTime)
            break
        else:
            print("Could not find time from " + server)

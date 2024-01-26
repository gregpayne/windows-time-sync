# TimeSync

A Time Sync script that can be run by Task Scheduler to keep the time of a device synced

## Python Script to Sync Time

The Python script in this repository can be executed as a Windows Scheduled Task to synchronise the PC time with a time server. The original file was downloaded from [here](https://gist.github.com/nihal111/23faa51c3f88a281b676dcbac77ce015). Once the script has been configured with the desired `server_list`, the script can be executed. This script will require Admin privilages to execute and when run directly the UAC dialog will show.

Once the script is working, set up a basic Task. The [following](https://www.esri.com/arcgis-blog/products/product/analytics/scheduling-a-python-script-or-model-to-run-at-a-prescribed-time/?rmedium=redirect&rsource=blogs.esri.com/esri/arcgis/2013/07/30/scheduling-a-scrip) article shows the process. Once the Task is configured, edit it to __run with highest priviledges__. This will allow the Task to run, without the UAC dialog getting in the way.

## Othere Options

These options are possible alternatives which have not yet been tested.

* timesyncweb has been downloaded from [here](https://answers.microsoft.com/en-us/windows/forum/all/how-to-force-windows-10-time-to-synch-with-a-time/20f3b546-af38-42fb-a2d0-d4df13cc8f43) and [here](https://www.robvanderwoude.com/sourcecode.php?src=timesyncweb_vbs)
* An application that could be used is [AboutTime](https://www.softpedia.com/get/Desktop-Enhancements/Clocks-Time-Management/AboutTime.shtml). I want to investigate a script solution first.
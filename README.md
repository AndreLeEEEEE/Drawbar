# Drawbar
Scrape data from the Production Requirements Planning or Workcenter Inventory Requirements page of Plex. Then, put that data onto
an excel sheet for importation to Monday.com

Versions of python and installed modules: 
- python 3.7.8
- selenium 3.141.0
- ChromeDriver 80.0.3987.106
- Visual Studio 16.8.6
- openpyxl 3.0.6
- numpy 1.19.5

Requirements:
Plex login for Wanco

When I asked one of the people who assigned me this project, "Are we using the Production Requirements Planning (PRP) method or
the Workcenter Inventory Requirements (WIR) method?" They responded with, "both of them." So this readme is gonna cover both
methods like the program. At the moment, the only functional method is PRP. The WIR method isn't even in the code yet, and it may
stay that way depending on what others require of this.

Update 11/23/2020: Since Kevin has confirmed the PRP method is doing what it's supposed to, there's no need to implement the WIR
method. The executable for this program exists and needs to be tested on another computer without the necessary construction tools
installed. Once physical restrictions are lifted, the executable needs to be demoed to Kit and someone else. I'm also still 
waiting on test Plex credentials so I can take my own out

Update 12/4/2020: I now have generic Plex credentials.

Update 12/18/2020: Back to my credentials since the generic account got nerfed in terms of what pages it can access.

Update 12/29/2020: Updated the column headers to match the new ones on the Material Request board.

Preliminary procedure:
The program utilizes the selenium module and the ChromeDriver to create a chrome web driver. This driver opens a new window for
Plex and logs in using the provided credentials. Once in, the driver will navigate to either the PRP page or WIR page.

PRP method:
On the PRP page, the script will change the following search criteria: the time frame window will be set to 1 week, the requirements
only box is checked, the suppress forecast box is checked, and the planning group is Drawbar Planning. Many results are likely to
appear. The program gets the total qty needed of a part, which is likely negative. Next, the program will click on the part name,
and then the Bill of Materials. From here, any component with a "-P" or "-E", which indicates paint or epoxy, will have their
number and pcs/pc recorded as they're drawbar components. Component pcs/pc will be multiplied by the qty of their overall part.
All duplicate components will be combined before all records are put into an excel sheet.

<p align="center">
<img width="200" src="icons/icon.png">
</p>

# Open Gantt
For a small private project I was looking for a very simple and free Gantt tool (especially with EXCEL export) but couldn't find one. There are lots of very good Gantt tools out there, but most of them are not free. Or offer a very limited feature set in their free accounts. So I decided to build Open Gantt. Currently it is still in Alpha/Beta phase. Just let me know your thoughts if you like.

### UI Description
Currently the tool offers the following features:

* Edit text/date bei clicking on the table cell
* **\<CTRL><Scroll\>** to zoom the Gantt View
* Context menu for adding/removing columns and lines 
* Move lines by dragging them
* Remembers last opened/saved file and re-opens it at app start

### Installation

#### Windows
Download and install the **.exe** installer.

#### Mac
Download and install the **.dmg** package.

#### Linux
For all common Linux distribuitions the **.AppImage** can be used without any installation (like a portable applications on Windows).

For Linux distributions using Debian packages, a Debian package is available as well. Just download the **.deb** package and install it as follows:
    
    sudo dpkg -i <package_name>.deb

If the command above fails, you might need to update the dependencies with:

    sudo apt install -f


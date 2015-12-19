License: This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 3 of the License, or (at your
option) any later version. This program is distributed in the hope that it
will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

To compile Tajpi correctly, you need the following:

To create just the program:
- Microsoft Visual Basic 6

To create the whole installer:
- Microsoft Visual Basic 6
- Inno Setup Compiler (http://www.innosetup.com/isinfo.php)
- Microsoft HTML Help Workshop (http://office.microsoft.com/en-us/orkXP/HA011362801033.aspx)

When compiling the installer you should create the various components in the following order:

1. Open Tajpi.vbp with Visual Basic 6 and compile Tajpi.exe.
2. Open Help\Esperanto\Helpo.hhp with HTML Help Workshop and compile Helpo.chm.
3. Open Help\English\Helpo (angla).hhp with HTML Help Workshop and compile Helpo (angla).chm.
4. Open Install.iss with Inno Setup Compiler, and compile the script. "Setup.exe" will be created in a directory named "Output", in the source code folder.


-----------------------------------------
Tajpi v2.97 - Klavarilo por esperantistoj
© 2008-2011 Thomas James
tmj2005@gmail.com
;Excel LDAP Search
;Copyright (C) 2008 Paul Brinkmann <paul@paulb.org>
;
;This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by 'the Free Software Foundation; either version 2 of the License, r (at your option) any later version.
;
;This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
;
;You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA




!macro IfFileInUse FILE GOTONOTLOCKED GOTOLOCKED

ClearErrors

Push $R0
  
  IfFileExists "${FILE}" 0 +3
	FileOpen $R0 "${FILE}" "a"
	FileClose $R0

Pop $R0

IfErrors "${GOTONOTLOCKED}" "${GOTOLOCKED}"

!macroend



#InstallKeybdHook
#SingleInstance force

Global KeyDownDelay = 30
Global SpaceDownDelay = 30

; Turn on/off removal of the numbers row in normal keyboard.
Global TurnOffNumberRow = 0

; -----------------------
; START: Left Blade Keys
; -----------------------

; Top row
addKey("q", "1")
addKey("w", "2")
addKey("e", "3")
addKey("r", "4")
addKey("t", "5")

; Middle row
addKey("a", "!")
addKey("s", "@", "~", "Shift", 0)
addKey("d", "#", "``")
addKey("f", "$")
addKey("g", "%")

; Bottom row
addKey("z", "+")
addKey("x", "-", "_")
addKey("c", "=", "|")
addKey("v", "{")
addKey("b", "}")

; -----------------------
; END  : Left Blade Keys
; -----------------------

; -----------------------
; START: Right Blade Keys
; -----------------------

; Top row
addKey("y", "6",, ["Control", "x"])
addKey("u", "7",, ["Control", "Left"],,, "Media_Prev")
addKey("i", "8",, "Up",,, "Volume_Up")
addKey("o", "9",, ["Control", "Right"],,, "Media_Next")
addKey("p", "0")

; Specials
addKey("[", "Delete",,"Backspace",,"Backspace")
addKey("]", "Delete",,"Backspace",,"Backspace")
addKey("\", "Delete",,"Backspace",,"Backspace")

; Middle row
addKey("h", "^",, ["Control", "c"])
addKey("j", "&",, "Left")
addKey("k", "*",, "Down",,, "Volume_Down")
addKey("l", "(",, "Right",,,"Media_Play_Pause")
addKey(";", ")",, ["Control", "z"],, "'")
addKey("'", "Enter",,,, "Enter")

; Bottom row
addKey("n", "[",, ["Control", "v"])
addKey("m", "]",, "Home",,, "Media_Prev")
addKey(",", ";",,,,, "Volume_Mute")
addKey(".", ":",, "End",,, "Media_Next")
addKey("/", "\",, ["Control", "y"])

; -----------------------
; END  : Right Blade Keys
; -----------------------
  
return

; Number row handler.
DoNumRowKeyDown(key) {
  if !TurnOffNumberRow {
    SendInput, {Blind}{%key% Down}{%key% Up}
  }
}

; Add key mappings function.
addKey(key, green_key, shift_green_key := -1, edit_key := -1, edit_key_up := 1, replace_key := -1, media_key := -1) {
  global keySet
  the_key := key
  if !isobject(keySet)
  {
    keySet := {}
  }
  if (replace_key != -1) {
    the_key := replace_key
  }

  newEntry := {key: the_key, green_key: green_key, has_shifted_green_key: !(shift_green_key = -1), shift_green_key: shift_green_key, has_edit_key: !(edit_key = -1), edit_key: edit_key, edit_key_up: (edit_key_up = 1), has_media_key: !(media_key = -1), media_key: media_key}
  keySet[key] := newEntry
}
    
; Get key mappings function.
getKey(key) {
  global keySet
  keys := keySet[key]
  return keys
}

; Add key mappings function.
setKeyDown(key, value) {
  global keydownSet
  if !isobject(keydownSet)
  {
    keydownSet := {}
  }
  keydownSet.Delete(key)
  newEntry := {isDown: value}
  keydownSet[key] := newEntry
}

; Get keydown
isKeyDown(key) {
  global keydownSet
  if !isobject(keydownSet)
  {
    keydownSet := {}
  }
  keydown := keydownSet[key]

  bval := keydown["isDown"]

  if %bval% {
    return 1
  }
  else {
    return 0
  }
}

isGreenLayer(delay := 1) {
  if GetKeyState("Space", "P") {
    if delay {
      ; Extra delay to make sure it is real.
      sleep, SpaceDownDelay ;
    }
    if GetKeyState("Space", "P") {
      return 1
    }
    else {
      return 0
    }
  }
  else {
    return 0
  }
}

isMediaLayer() {
  if GetKeyState("d", "P") and GetKeyState("f", "P") and GetKeyState("s", "P") and GetKeyState("a", "P")
  {
    return 1
  }
  else {
    return 0
  }
}

isEditLayer() {
  if GetKeyState("d", "P") and GetKeyState("f", "P")
  {
    return 1
  }
  else {
    return 0
  }
}

DoKeyDown(key) {
  global eat_space

  keyobject := getKey(key)
  send_key := keyobject["key"]

  if !isKeyDown(key) {
    sleep, %KeyDownDelay% ; Delay before sending key.
  }
  if isGreenLayer() {
    eat_space := 1
    gkey := keyobject["green_key"]
    if GetKeyState("Shift", "P") {
      if keyobject["has_shifted_green_key"] {
        gkey := keyobject["shift_green_key"]
      }
      SendInput, {Blind}{Shift Up}{%gkey% down}{%gkey% up}{Shift Down}
    }
    else {
      SendInput, {Blind}{%gkey% down}{%gkey% up}
    }
  }
  else {
    if isMediaLayer() {
      if keyobject["has_media_key"] {
	media_key := keyobject["media_key"]
	if keyobject["media_key_up"] {
	  if IsObject(media_key) {
	    len := media_key.Length() + 1
	    Loop % media_key.Length() {
	      tmpkey := media_key[A_Index]
	      SendInput, {Blind}{%tmpkey% Down}
	    }
	    Loop % media_key.Length() {
	      tmpkey := media_key[len - A_Index]
	      SendInput, {Blind}{%tmpkey% Up}
	    }
          }
	  else {
            SendInput, {Blind}{%media_key% down}{Blind}{%media_key% up}
	  }
        }
	else {
          SendInput, {Blind}{%media_key% down}
	}
      }
      else {
	; Do nothing.
      }
    }
    else {
      if isEditLayer() {
        if keyobject["has_edit_key"] {
  	edit_key := keyobject["edit_key"]
  	if keyobject["edit_key_up"] {
  	  if IsObject(edit_key) {
  	    len := edit_key.Length() + 1
  	    Loop % edit_key.Length() {
  	      tmpkey := edit_key[A_Index]
  	      SendInput, {Blind}{%tmpkey% Down}
  	    }
  	    Loop % edit_key.Length() {
  	      tmpkey := edit_key[len - A_Index]
  	      SendInput, {Blind}{%tmpkey% Up}
  	    }
            }
  	  else {
              SendInput, {Blind}{%edit_key% down}{Blind}{%edit_key% up}
  	  }
          }
  	else {
            SendInput, {Blind}{%edit_key% down}
  	}
        }
        else {
  	; Do nothing.
        }
      }
      else {
        SendInput, {Blind}{%send_key% down}{%send_key% up}
      }
    }
  }

  if GetKeyState(key, "P")
    setKeyDown(key, 1)
}

DoKeyUp(key) {
  setKeyDown(key, 0)
}

#MaxThreadsBuffer On
; -------------------------
; START: Left Blade Hotkeys
; -------------------------

; Top row

*q::
  DoKeyDown("q")
  return

*q up::
  DoKeyUp("q")
  return

*w::
  DoKeyDown("w")
  return

*w up::
  DoKeyUp("w")
  return

*e::
  DoKeyDown("e")
  return

*e up::
  DoKeyUp("e")
  return

*r::
  DoKeyDown("r")
  return

*r up::
  DoKeyUp("r")
  return

*t::
  DoKeyDown("t")
  return

*t up::
  DoKeyUp("t")
  return

; Middle row

*a::
  DoKeyDown("a")
  return

*a up::
  DoKeyUp("a")
  return

*s::
  DoKeyDown("s")
  return

*s up::
  DoKeyUp("s")
  if GetKeyState("Shift", "D") and !(GetKeyState("Shift", "P")) {
    SendInput, {Blind}{Shift Up}
  }
  return

*d::
  DoKeyDown("d")
  return

*d up::
  DoKeyUp("d")
  return

*f::
  DoKeyDown("f")
  return

*f up::
  DoKeyUp("f")
  return

*g::
  DoKeyDown("g")
  return

*g up::
  DoKeyUp("g")
  return


; Bottom row

*z::
  DoKeyDown("z")
  return

*z up::
  DoKeyUp("z")
  return

*x::
  DoKeyDown("x")
  return

*x up::
  DoKeyUp("x")
  return

*c::
  DoKeyDown("c")
  return

*c up::
  DoKeyUp("c")
  return

*v::
  DoKeyDown("v")
  return

*v up::
  DoKeyUp("v")
  return

*b::
  DoKeyDown("b")
  return

*b up::
  DoKeyUp("b")
  return

; -------------------------
; END  : Left Blade Hotkeys
; -------------------------

; --------------------------
; START: Right Blade Hotkeys
; --------------------------

; Top row

*y::
  DoKeyDown("y")
  return

*y up::
  DoKeyUp("y")
  return

*u::
  DoKeyDown("u")
  return

*u up::
  DoKeyUp("u")
  return

*i::
  DoKeyDown("i")
  return

*i up::
  DoKeyUp("i")
  return

*o::
  DoKeyDown("o")
  return

*o up::
  DoKeyUp("o")
  return

*p::
  DoKeyDown("p")
  return

*p up::
  DoKeyUp("p")
  return

*[::
  DoKeyDown("[")
  return
  
*[ up::
   DoKeyUp("[")
   return
     
*]::
  DoKeyDown("]")
  return

*] up::
  DoKeyUp("]")
  return

*\::
  DoKeyDown("\")
  return

*\ up::
  DoKeyUp("\")
  return



; Middle row

*h::
  DoKeyDown("h")
  return

*h up::
  DoKeyUp("h")
  return

*j::
  DoKeyDown("j")
  return

*j up::
  DoKeyUp("j")
  return

*k::
  DoKeyDown("k")
  return

*k up::
  DoKeyUp("k")
  return

*l::
  DoKeyDown("l")
  return

*l up::
  DoKeyUp("l")
  return

*;::
  DoKeyDown(";")
  return

*; up::
  DoKeyUp(";")
  return

*'::
  DoKeyDown("'")
  return

*' up::
  DoKeyUp("'")
  return

; Bottom row

*n::
  DoKeyDown("n")
  return

*n up::
  DoKeyUp("n")
  return

*m::
  DoKeyDown("m")
  return

*m up::
  DoKeyUp("m")
  return

*,::
  DoKeyDown(",")
  return
  
*, up::
  DoKeyUp(",")
  return
  
*.::
  DoKeyDown(".")
  return
  
*. up::
  DoKeyUp(".")
  return
  
*/::
  DoKeyDown("/")
  return
  
*/ up::
  DoKeyUp("/")
  return

; --------------------------
; END  : Right Blade Hotkeys
; --------------------------

; --------------------------
; START: Space Blade Hotkeys
; --------------------------

*Space::
  Global eat_space
  eat_space := 0
  return

*Space up::
  Global eat_space
  if !eat_space {
    SendInput, {Blind}{Space Down}{Space Up}
  }
  eat_space := 0
  return

; --------------------------
; END  : Space Blade Hotkeys
; --------------------------

; --------------------------
; START: Tab handling
; --------------------------

!Tab::SendInput, {Blind}{Tab Down}{Tab Up}
!Tab Up::return

*Tab::
  Global eat_tab
  eat_tab := 0
  return

*Tab up::
  Global eat_tab

  if isGreenLayer(0) {
    eat_tab := 1
    eat_space := 1
    SendInput, {Blind}{Esc Down}{Blind}{Esc Up}
  }

  if !eat_tab {
    SendInput, {Blind}{Tab Down}{Blind}{Tab Up}
  }
  eat_tab := 0
  return

; --------------------------
; END  : Tab handling
; --------------------------

; -----------------------
; START: Number row
; -----------------------

*`::
  DoNumRowKeyDown("``")
  return
*` up::return
*1::
  DoNumRowKeyDown("1")
  return
*1 up::return
*2::
  DoNumRowKeyDown("2")
  return
*2 up::return
*3::
  DoNumRowKeyDown("3")
  return
*3 up::return
*4::
  DoNumRowKeyDown("4")
  return
*4 up::return
*5::
  DoNumRowKeyDown("5")
  return
*5 up::return
*6::
  DoNumRowKeyDown("6")
  return
*6 up::return
*7::
  DoNumRowKeyDown("7")
  return
*7 up::return
*8::
  DoNumRowKeyDown("8")
  return
*8 up::return
*9::
  DoNumRowKeyDown("9")
  return
*9 up::return
*0::
  DoNumRowKeyDown("0")
  return
*0 up::return
*-::
  DoNumRowKeyDown("-")
  return
*- up::return
*+::
  DoNumRowKeyDown("+")
  return
*+ up::return
    
; -----------------------
; END : Number row
; -----------------------

#InstallKeybdHook
#SingleInstance force

Global KeyDownDelay = 40
Global SpaceDownDelay = 60

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
addKey("s", "@", 1, "~")
addKey("d", "#", 1, "``")
addKey("f", "$")
addKey("g", "$")

; Bottom row
addKey("z", "+")
addKey("x", "-", 1, "_")
addKey("c", "=", 1, "|")
addKey("v", "{")
addKey("b", "}")

; -----------------------
; END  : Left Blade Keys
; -----------------------

; -----------------------
; START: Right Blade Keys
; -----------------------

; Top row
addKey("y", "6")
addKey("u", "7")
addKey("i", "8")
addKey("o", "9")
addKey("p", "0")

; Middle row
addKey("h", "^")
addKey("j", "&")
addKey("k", "*")
addKey("l", "(")
addKey(";", ")", 1, """")

; Bottom row
addKey("n", "+")
addKey("m", "-")
addKey(",", "=", 1, "<")
addKey(".", "{", 1, ">")
addKey("/", "}", 1, "?")

; -----------------------
; END  : Right Blade Keys
; -----------------------
  
return

; Add key mappings function.
addKey(key, green_key, has_shifted_green_key := 0, shift_green_key := 0) {
  global keySet
  if !isobject(keySet)
  {
    keySet := {}
  }
  newEntry := {green_key: green_key, has_shifted_green_key: has_shifted_green_key, shift_green_key: shift_green_key}
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

DoKeyDown(key) {
  global eat_space

  keyobject := getKey(key)

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
      Send, {Blind}{Shift Up}{%gkey% down}{%gkey% up}{Shift Down}
    }
    else {
      Send, {Blind}{%gkey% down}{%gkey% up}
    }
  }
  else {
    Send, {Blind}{%key% down}{%key% up}
  }

  if GetKeyState(key, "P")
    setKeyDown(key, 1)
}

DoModKeyDown(key) {
  global eat_space

  keyobject := getKey(key)

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
      Send, {Blind}{Shift Up}{%gkey% down}{%gkey% up}{Shift Down}
    }
    else {
      Send, {Blind}{%gkey% down}{%gkey% up}
    }
  }
  else {
    Send, {Blind}{%key% down}{%key% up}
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
    Send, {Blind}{Space Down}{Space Up}
  }
  eat_space := 0
  return

*Tab::
  Global eat_tab
  eat_tab := 0
  return

*Tab up::
  Global eat_tab

  if isGreenLayer(0) {
    eat_tab := 1
    eat_space := 1
    Send, {Blind}{Esc Down}{Esc Up}
  }

  if !eat_tab {
    Send, {Blind}{Tab Down}{Tab Up}
  }
  eat_tab := 0
  return

; --------------------------
; END  : Space Blade Hotkeys
; --------------------------

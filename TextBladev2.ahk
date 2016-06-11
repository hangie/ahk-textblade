#InstallKeybdHook
#SingleInstance force

Global TB_KeyDelay := 300
Global TB_Layers := {}

SetStoreCapslockMode, off

; Toying with preventing delays to make sure correct keys output.
SetBatchLines, -1

Global xls
if not(xls)
{
  xls := ComObjCreate("Excel.Application")
  Path := A_ScriptDir . "\TextBlade.xlsx"
  wb := xls.Workbooks.Open(Path)
  xls.Visible := False
}


For Key in xls.Worksheets("Keys").Range("Hotkeys[Hotkey]")
{
  value := Key.Value2

  ; Add the key and it's values.
  addKey(value, Key)

  ; Just the key
  Hotkey, %value%, DoHotKeyDown
  Hotkey, %value% Up, DoHotKeyUp
}

wb.Close(false)
xls.Quit

return
#MaxThreadsBuffer On

MySleep(period)
{
  DllCall("Sleep", Uint, period)
}

; Get keydown
isKeyDown(key) {
  global keydownSet
  if !isobject(keydownSet)
  {
    keydownSet := {}
  }
  keydown := keydownSet[key]
  if (isobject(keydown))
  {
    return 1
  }
  else
  {
    return 0
  }
}

; Add key mappings function.
setKeyDown(key, value) {
  global keydownSet
  if !isobject(keydownSet)
  {
    keydownSet := {}
  }
  keydownSet.Delete(key)

  if (value != 0)
  {
    newEntry := {isDown: value}
    keydownSet[key] := newEntry
  }
}

isGreenLayer(delay := 1) {
  global eat_space
  if GetKeyState("Space", "P") {
    if delay {
      ; Extra delay to make sure it is real.
      MySleep(SpaceDownDelay)
    }
    if GetKeyState("Space", "P") {
      return 1
    }
    else {
      if isKeyDown("Space")
      {
        eat_space := 1
        SendInput, {Space}
        setKeyDown("Space", 0)
      }
      return 0
    }
  }
  else {
    if isKeyDown("Space")
    {
      SendInput, {Space}
      eat_space := 1
      setKeyDown("Space", 0)
    }
    return 0
  }
}

; Add key mappings function.
addKey(key, row) {
  global keySet
  global TB_Layers
  the_key := key
  if !isobject(keySet)
  {
    keySet := {}
  }

  newEntry := {key: key, down_key: row.Offset(0, 1).Value2, up_key: row.Offset(0, 2).Value2, green_down_key: row.Offset(0, 3).Value2, green_up_key: row.Offset(0, 4).Value2}
  ; Check for layer keys
  layer := row.Offset(0, 5).Value2
  if (StrLen(layer) > 0)
  {
    layer_keys_str := row.Offset(0, 6).Value2
    layer_keys := {}
    Loop, Parse, layer_keys_str, `n
    {
      layer_keys[A_LoopField] := 1
    }
    newEntry["layer"] := {layer: layer, keys: layer_keys}
    l := newEntry["layer"]
    if (isObject(l))
    {
      lname := l["layer"]
      TB_Layers[lname] := 0
    }
  }
  keySet[key] := newEntry
}

; Get key mappings function.
getKey(key) {
  global keySet
  keys := keySet[key]
  return keys
}

getLayer(key)
{
  the_key := getKey(key)
  layer := the_key["layer"]
  return layer
}

getLayerKeys(key)
{
  layer := key["layer"]
  if (!isObject(layer))
  {
    return 0
  }

  layer_keys := layer["keys"]
  return layer_keys
}

getDownKey(key) {
  keys := getKey(key)
  the_key := keys["down_key"]
  return the_key
}

getUpKey(key) {
  keys := getKey(key)
  the_key := keys["up_key"]
  return the_key
}

getGreenDownKey(key) {
  keys := getKey(key)
  the_key := keys["green_down_key"]
  return the_key
}

getGreenUpKey(key) {
  keys := getKey(key)
  the_key := keys["green_up_key"]
  return the_key
}

DoHotKeyDown:
  Critical, 1000
  key := ""
  if (!isKeyDown(A_ThisHotKey))
  {
    setKeyDown(A_ThisHotKey, 1)
    MySleep(TB_KeyDelay)
  }
  if (isGreenLayer())
  {
    key := getGreenDownKey(A_ThisHotKey)
    layer := "Green"
    setKeyDown("Space", 0)
  }
  else
  {
    layer := getLayer(A_ThisHotKey)
    if (isObject(layer))
    {
      layer_name := layer["layer"]
      in_layer := 1
      for lkey in layer["keys"]
      {
        if (!GetKeyState(lkey, "P"))
        {
          in_layer := 0
          break
        }
      }
      if (in_layer)
      {
        TB_Layers[layer_name] := 1
      }
      else
      {
        key := getDownKey(A_ThisHotKey)
      }
    }
    else
    {
      key := getDownKey(A_ThisHotKey)
    }
  }
  SendInput, %key%
  return

DoHotKeyUp:
  Critical, 1000
  StringLeft, this_key, A_ThisHotKey, StrLen(A_ThisHotKey) - 3

  if (isKeyDown(this_key))
  {
    key := getUpKey(this_key)
    SendInput, %key%
    setKeyDown(this_key, 0)
  }
  return

OnExit:
  xls.Quit
  return

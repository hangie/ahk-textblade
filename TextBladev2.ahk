#InstallKeybdHook
#SingleInstance force

Global TB_KeyDelay := 100
Global TB_GreenKeyDelay := 100
Global TB_Layers := {}
Global TB_LayerKeys := {}
Global TB_Layer := {}

TB_Layer.push("Alpha")

SetStoreCapslockMode, off

; Toying with preventing delays to make sure correct keys output.
SetBatchLines, -1

Global xls

; Open Excel
if not(xls)
{
  xls := ComObjCreate("Excel.Application")
  Path := A_ScriptDir . "\TextBlade.xlsx"
  wb := xls.Workbooks.Open(Path)
  xls.Visible := False
}


; Load hotkeys.
For Key in xls.Worksheets("Keys").Range("Hotkeys[Hotkey]")
{
  value := Key.Value2

  ; Loop over header.
  key_entry := {key: value}
  For Col in xls.Worksheets("Keys").Range("HotKeys[#Headers]")
  {
    heading := Col.Value2
    col_value := Col.Offset(Key.Row - 1 , 0).Value2
    key_entry[heading] := col_value
  }
  ; MsgBox % key_entry["Hotkey"]
  addKeyEntry(value, key_entry)

  ; Just the key
  Hotkey, %value%, DoHotKeyDown
  Hotkey, %value% Up, DoHotKeyUp
}

; Load layers
For Layer in xls.Worksheets("Layers").Range("Layers[Layer]")
{
  layer_name := Layer.Value2
  layer_keys_str := Layer.Offset(0, 1).Value2

  ; Split keys.
  layer_keys := {}
  Loop, Parse, layer_keys_str, `n
  {
    layer_keys[A_LoopField] := 1
    key_layers := TB_LayerKeys[A_LoopField]
    if (!isObject(key_layers))
    {
      key_layers := {}
    }
    key_layers[layer_name] := 1
    TB_LayerKeys[A_LoopField] := key_layers
  }
  TB_Layers[layer_name] := layer_keys
}

; Close Excel.
wb.Close(false)
xls.Quit

isLayerComplete("f")

return


#MaxThreadsBuffer On

isLayerKey(key)
{
  Global TB_LayerKeys
  lkey := TB_LayerKeys[key]
  MsgBox % key " => " isObject(lkey)
  return isObject(lkey)
}

isLayerComplete(key)
{
  Global TB_LayerKeys
  Global TB_Layers
  Global TB_Layer
  if (isLayerKey(key))
  {
    layers := TB_LayerKeys[key]
    For layer_name in layers
    {
      MsgBox % key " => " layer_name
      layer := TB_Layers[layer_name]
      complete := 1
      For lkey in layer
      {
        if (!GetKeyState(lkey, "P"))
        {
          complete := 0
          break
        }
      }
      if (complete)
      {
        TB_Layer := layer_name
        MsgBox % key " => complete[" complete "]"
        TB_Layer.push(layer_name)
      }
    }
  }
  else
  {
    return 0
  }
}

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

; Get keydown
getKeyDown(key) {
  global keydownSet
  if !isobject(keydownSet)
  {
    keydownSet := {}
  }
  keydown := keydownSet[key]
  return keydown
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
    newEntry := {layer: value}
    keydownSet[key] := newEntry
  }
}

deleteKeyDown(key) {
  global keydownSet
  if !isobject(keydownSet)
  {
    keydownSet := {}
  }
  keydownSet.Delete(key)
}

isGreenLayer(delay := 1) {
  global eat_space
  if GetKeyState("Space", "P") {
    if delay {
      ; Extra delay to make sure it is real.
      MySleep(TB_GreenKeyDelay)
    }
    if GetKeyState("Space", "P") {
      return 1
    }
    else {
      if isKeyDown("Space")
      {
        eat_space := 1
        SendInput, {Space}
        setKeyDown("Space", "Green")
      }
      return 0
    }
  }
  else {
    if isKeyDown("Space")
    {
      SendInput, {Space}
      eat_space := 1
      setKeyDown("Space", "Green")
    }
    return 0
  }
}

addKeyEntry(key, newEntry) {
  global keySet
  the_key := key
  if !isobject(keySet)
  {
    keySet := {}
  }

  keySet[key] := newEntry
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

  ; Retrieve the layer for the key.
  key_down := getKeyDown(key)
  layer := "Alpha"
  if (isObject(key_down))
  {
    layer := key_down["layer"]
    if (StrLen(layer) <= 0)
    {
      layer := "Alpha"
    }
  }

  key_down := layer " Down"
  
  the_key := keys[key_down]
  ; MsgBox % key_down " => " the_key
  return the_key
}

getUpKey(key) {
  keys := getKey(key)

  ; Retrieve the layer for the key.
  key_down := getKeyDown(key)
  layer := "Alpha"
  if (isObject(key_down))
  {
    layer := key_down["layer"]
    if (StrLen(layer) <= 0)
    {
      layer := "Alpha"
    }
  }

  key_up := layer " Up"
  
  the_key := keys[key_up]
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

getCurrentLayer()
{
  layer := "Alpha"
  if (isGreenLayer())
  {
    layer := "Green"
  }
  return layer
}

DoHotKeyDown:
  Critical, 1000
  key := ""
  if (!isKeyDown(A_ThisHotKey))
  {
    MySleep(TB_KeyDelay)
    isLayerComplete(A_ThisHotKey)
    setKeyDown(A_ThisHotKey, getCurrentLayer())
  }
  if (isGreenLayer(0))
  {
    key := getDownKey(A_ThisHotKey)
    layer := "Green"
    setKeyDown("Space", "Green")
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

  key := getUpKey(this_key)
  SendInput, %key%
  deleteKeyDown(this_key)

  return

OnExit:
  xls.Quit
  return

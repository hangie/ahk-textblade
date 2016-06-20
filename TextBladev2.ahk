#InstallKeybdHook
#SingleInstance force

Global TB_KeyDelay := 40
Global TB_Layers := {}
Global TB_LayerKeys := {}
Global TB_Layer := {}
Global TB_ActiveLayers := {}
Global TB_LayerDelay := {}
Global TB_Layout

IniRead, TB_Layout, TextBladev2.ini, Config, Layout, "QWERTY"

Global TB_LayoutSheet := "Keys." . TB_Layout
Global TB_LayoutTable := "HotKeys." . TB_Layout

TB_Layer := "Alpha"

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
Critical 10000
For Key in xls.Worksheets(TB_LayoutSheet).Range(TB_LayoutTable . "[Hotkey]")
{
  value := Key.Value2

  ; Loop over header.
  key_entry := {key: value}
  For Col in xls.Worksheets(TB_LayoutSheet).Range(TB_LayoutTable . "[#Headers]")
  {
    heading := Col.Value2
    col_value := Col.Offset(Key.Row - 1 , 0).Value2
    key_entry[heading] := col_value
  }
  Col := ""

  ; MsgBox % key_entry["Hotkey"]
  addKeyEntry(value, key_entry)

  ; Just the key
  Hotkey, %value%, DoHotKeyDown
  Hotkey, %value% Up, DoHotKeyUp
}
Key := ""

; Load layers
For Layer in xls.Worksheets("Layers").Range("Layers[Layer]")
{
  layer_name := Layer.Value2
  layer_keys_str := Layer.Offset(0, 1).Value2
  layer_delay := Layer.Offset(0, 2).Value2

  ; Split keys.
  layer_keys := {}
  Loop, Parse, layer_keys_str, `n
  {
    layer_keys[A_LoopField] := 1

    ; Retrieve the current layers for this key.
    key_layers := TB_LayerKeys[A_LoopField]
    if (!isObject(key_layers))
    {
      ; No layers yet.  Initialize.
      key_layers := {}
    }
    ; Add this layer as a layer for this key.
    key_layers[layer_name] := 1
    TB_LayerKeys[A_LoopField] := key_layers
  }
  ; Set the keys for this layer.
  TB_Layers[layer_name] := layer_keys
  TB_LayerDelay[layer_name] := layer_delay
}
Layer := ""

; Close Excel.
wb.Close(false)
xls.Quit
wb := ""
xls := ""

; MsgBox "Test if layer complete"
if (isLayerComplete("f"))
{
  MsgBox % "f key has complete layer"
}

return


#MaxThreadsBuffer On

isLayerKey(key)
{
  Global TB_LayerKeys
  lkey := TB_LayerKeys[key]
  ; MsgBox % key " => " isObject(lkey)
  return isObject(lkey)
}

isLayerComplete(key)
{
  Global TB_LayerKeys
  Global TB_Layers
  Global TB_Layer
  Global TB_ActiveLayers
  if (isLayerKey(key))
  {
    ; Delay one key length.
    ; MySleep(TB_KeyDelay)
    clayers := ""

    ; Get the layers for the key.
    layers := TB_LayerKeys[key]
    tmp_layer := ""
    tmp_layer_size := 0

    For layer_name in layers
    {
      layer_delay := TB_LayerDelay[layer_name]
      if (layer_delay > 0)
      {
        MySleep(layer_delay)
      }

      layer := TB_Layers[layer_name]
      complete := 1
      lsize := 0
      ; MsgBox % layer_name " has layer " layer
      For lkey in layer
      {
        lsize := lsize + 1
        ; MsgBox % layer_name " has key " lkey
        if (!GetKeyState(lkey, "P"))
        {
          ; MsgBox % layer_name " is not complete as " lkey " is not down"
          complete := 0
        }
        else if (isKeyDown(lkey))
        {
          ; Key already processed.  Ignore
          ; MsgBox % layer_name " is not complete as " lkey " is already processed"
          complete := 0
        }
      }
      if (complete)
      {
        ; Add layer to active layers.
        TB_ActiveLayers[layer_name] := lsize
        clayers := clayers . " " layer_name

        ; MsgBox % layer_name " is complete"
        ; Layer with most keys is primary active layer.
        if (lsize > tmp_layer_size)
        {
          tmp_layer_size := lsize
          tmp_layer := layer_name
        }
      }
    }
    if (tmp_layer_size > 0)
    {
      ; Set active layer.
      TB_Layer := tmp_layer
      ; MsgBox % key " => complete[" tmp_layer "]" clayers

      ; Mark the keys of the layer as down.
      layer := TB_Layers[tmp_layer]
      For lkey in layer
      {
        setKeyDown(lkey, TB_Layer)
      }

      return 1
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
    ; MsgBox % key " is " keydown
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
    ; MsgBox % key "[" value "] Down"
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

DoHotKeyDown:
  Critical, 1000
  Global TB_KeyDelay
  key := ""
  hkey := A_ThisHotKey
  if (!isKeyDown(hkey))
  {
    k := getKey(hkey)
    kdelay := k["Delay"]
    if (!kdelay)
    {
      kdelay := TB_KeyDelay
    }
    ; KeyWait, %hkey%, T%kdelay%
    MySleep(kdelay)
    if (!isLayerComplete(hkey))
    {
      setKeyDown(hkey, TB_Layer)
    }
  }
  else
  {
    ; MsgBox % "THERE"
  }
  key := getDownKey(hkey)

  SendInput, %key%
  return

DoHotKeyUp:
  Critical, 1000
  Global TB_LayerKeys
  Global TB_ActiveLayers
  StringLeft, this_key, A_ThisHotKey, StrLen(A_ThisHotKey) - 3

  key := getUpKey(this_key)

  SendInput, %key%
  ; MsgBox % "Deleting key " this_key
  deleteKeyDown(this_key)

  ; Remove active layers for key.
  if (isLayerKey(this_key))
  {
    ; Delete active layers for this key.
    layers := TB_LayerKeys[this_key]
    For layer_name in layers
    {
      TB_ActiveLayers.Delete(layer_name)
    }

    ; Find which layer is now active.
    active_layer := "Alpha"
    lkey_size := 0
    For layer in TB_ActiveLayers
    {
      if (TB_ActiveLayers[layer] > lkey_size)
      {
        lkey_size := TB_ActiveLayers[layer]
        active_layer := layer
      }
    }

    ; Set the active layer.
    TB_Layer := active_layer
  }
  return

OnExit:
  xls.Quit
  return

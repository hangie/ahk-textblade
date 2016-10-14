#InstallKeybdHook
#SingleInstance force

Global TB_KeyDelay := 40
Global TB_Layers := {}
Global TB_LayerKeys := {}
Global TB_Layer := {}
Global TB_ActiveLayers := {}
Global TB_LayerDelay := {}
Global TB_LayerAliases := {}
Global TB_Layout

Global TB_Modifiers := {}
Global TB_ModifierKeys := {}
Global TB_ModifierDelay := {}
Global TB_Modifier := {}
Global TB_ActiveModifiers := {}
Global TB_ModKeysDown := ""
Global TB_ModKeysUp := ""

IniRead, TB_Layout, TextBladev2.ini, Config, Layout, "QWERTY"
IniRead, TB_KeyDelay, TextBladev2.ini, Config, KeyDelay, 40

Global TB_LayoutSheet := "Keys." . TB_Layout
Global TB_LayoutTable := "HotKeys." . TB_Layout

TB_Layer := "Alpha"
TB_Modifier := ""

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

For LayerAlias in xls.Worksheets("Layer Aliases").Range("LayerAliases[LayerAlias]")
{
  Global TB_LayerAliases
  
  layer_alias := LayerAlias.Value2
  layer_name := LayerAlias.Offset(0, 1).Value2
  TB_LayerAliases[layer_alias] := layer_name
}
LayerAlias := ""

; Load modkeys
For Modifier in xls.Worksheets("Modifiers").Range("Modifiers[Modifier]")
{
  modifier_name := Modifier.Value2
  modifier_keys_str := Modifier.Offset(0, 1).Value2
  modifier_delay := Modifier.Offset(0, 2).Value2
  down_keys := Modifier.Offset(0, 3).Value2
  up_keys := Modifier.Offset(0, 4).Value2

  ; MsgBox % modifier_name " has keys " modifier_keys_str

  ; Split keys.
  modifier_keys := {}
  Loop, Parse, modifier_keys_str, `n
  {
    modifier_keys[A_LoopField] := 1

    ; Retrieve the current modifiers for this key.
    key_modifiers := TB_ModifierKeys[A_LoopField]
    if (!isObject(key_modifiers))
    {
      ; No modifiers yet.  Initialize.
      key_modifiers := {}
    }
    ; Add this modifier as a modifier for this key.
    key_modifiers[modifier_name] := 1
    TB_ModifierKeys[A_LoopField] := key_modifiers
  }
  ; Set the keys for this modifier.
  TB_Modifiers[modifier_name] := {keys: modifier_keys, delay: modifier_delay, key_down: down_keys, key_up: up_keys}
  TB_ModifierDelay[modifier_name] := modifier_delay
}
Modifier := ""

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
  if (isLayerKey(value) or isModifierKey(value))
  {
    isSpace := RegExMatch(value, "Space", theKeys)
    if (isSpace)
    {
      Hotkey, %value%, DoSpaceHotKeyDown
      Hotkey, %value% Up, DoSpaceHotKeyUp
    }
    else
    {
      Hotkey, %value%, DoComplexHotKeyDown
      Hotkey, %value% Up, DoComplexHotKeyUp
    }
  }
  else
  {
    Hotkey, %value%, DoHotKeyDown
    Hotkey, %value% Up, DoHotKeyUp
  }
}
Key := ""

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

; MsgBox "Test modifiers"
if (isModifierKey("c"))
{
  ; MsgBox % "c key is a modifier"
}
else
{
  ; MsgBox % "c key is NOT a modifier"
}

; MsgBox "Test layers"
if (isLayerKey("d"))
{
  ; MsgBox % "d key is a layer"
}
else
{
  ; MsgBox % "d key is NOT a layer"
}

GetModKeyState("+Space")

return


#MaxThreadsBuffer On

isLayerKey(key)
{
  Global TB_LayerKeys
  lkey := TB_LayerKeys[key]
  ; MsgBox % key " => " isObject(lkey)
  return isObject(lkey)
}

isModifierKey(key)
{
  Global TB_ModifierKeys
  lkey := TB_ModifierKeys[key]
  return isObject(lkey)
}

isActiveModifierKey(key)
{
  Global TB_ModifierKeys
  Global TB_ActiveModifiers
  ret_val := 0

  if (isModifierKey(key))
  {
    key_modifiers := TB_ModifierKeys[key]
    if (isobject(key_modifiers))
    {
      For modifier in key_modifiers
      {
        if (TB_ActiveModifiers[modifier] > 0)
        {
          ret_val := 1
        }
      }
    }
  }

  return ret_val
}

GetModKeyState(keys, mode := "P")
{
  hasModKeys := RegExMatch(keys, "^([\^\+#!]+)", ModKeys)
  if (hasModKeys)
  {
    StringRight, RestKeys, keys, StrLen(keys) - StrLen(ModKeys)

    if (InStr(ModKeys, "+"))
    {
      if (!GetKeyState("Shift", mode))
      {
        return 0
      }
    }
    if (InStr(ModKeys, "^"))
    {
      if (!GetKeyState("Control", mode))
      {
        return 0
      }
    }
    if (InStr(ModKeys, "!"))
    {
      if (!GetKeyState("Alt", mode))
      {
        return 0
      }
    }
    if (InStr(ModKeys, "#"))
    {
      if (!GetKeyState("Win", mode))
      {
        return 0
      }
    }

    return GetKeyState(RestKeys, mode)
  }
  else
  {
    return GetKeyState(keys, mode)
  }
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
        if (!GetModKeyState(lkey, "P"))
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

isComplexComplete(key)
{
  Global TB_ModifierKeys
  Global TB_Modifiers
  Global TB_Modifier
  Global TB_ActiveModifiers
  Global TB_LayerKeys
  Global TB_Layers
  Global TB_Layer
  Global TB_ActiveLayers
  Global TB_ModKeysDown
  Global TB_ModKeysUp

  ret_val := 0

  is_mod_key := isModifierKey(key)
  is_layer_key := isLayerKey(key)
  
  if (is_mod_key or is_layer_key)ccd
  {
    ; Delay one key length.
    ; MySleep(TB_KeyDelay)
    cmodifiers := ""

    ; Get the modifiers and layers for the key.
    modifiers := TB_ModifierKeys[key]
    layers := TB_LayerKeys[key]
    tmp_modifier := ""
    tmp_modifier_size := 0
    tmp_layer := ""
    tmp_layer_size := 0

    ; Only delay once.  Find maximum delay.
    complex_delay := 0
    For modifier_name in modifiers
    {
      modifier_delay := TB_ModifierDelay[modifier_name]
      if (modifier_delay > complex_delay)
      {
        complex_delay := modifier_delay
      }
    }
    For layer_name in layers
    {
      layer_delay := TB_LayerDelay[layer_name]
      if (layer_delay > complex_delay)
      {
        complex_delay := layer_delay
      }
    }

    ; Do delay
    if (complex_delay > 0)
    {
      MySleep(complex_delay)
    }

    ; Process modifiers
    For modifier_name in modifiers
    {
      modifier := TB_Modifiers[modifier_name]
      complete := 1
      lsize := 0
      ; MsgBox % modifier_name " has modifier " modifier
      For lkey in modifier.keys
      {
        lsize := lsize + 1
        ; MsgBox % modifier_name " has key " lkey
        if (!GetModKeyState(lkey, "P"))
        {
          ; MsgBox % modifier_name " is not complete as " lkey " is not down"
          complete := 0
        }
        else if (isKeyDown(lkey))
        {
          ; Key already processed.  Ignore
          ; MsgBox % modifier_name " is not complete as " lkey " is already processed"
          complete := 0
        }
      }
      if (complete)
      {
        ; Add modifier to active modifiers.
        if (!isobject(TB_ActiveModifiers[modifier_name]))
        {
          TB_ActiveModifiers[modifier_name] := lsize
          down_key := TB_Modifiers[modifier_name].key_down
        }
        cmodifiers := cmodifiers . " " modifier_name

        ; MsgBox % modifier_name " is complete"
        ; Modifier with most keys is primary active modifier.
        if (lsize > tmp_modifier_size)
        {
          tmp_modifier_size := lsize
          tmp_modifier := modifier_name
        }
      }
    }
    if (tmp_modifier_size > 0)
    {
      ; Set active modifier.
      TB_Modifier := tmp_modifier

      TB_ModKeysDown := ""
      TB_ModKeysUp := ""
      For modifier in TB_ActiveModifiers
      {
        TB_ModKeysDown := TB_ModKeysDown TB_Modifiers[modifier].key_down
        TB_ModKeysUp := TB_ModKeysUp TB_Modifiers[modifier].key_up
      }
      ; MsgBox % "Mod keys are " TB_ModKeysDown 

      ; Mark the keys of the modifier as down.
      modifier := TB_Modifiers[tmp_modifier]
      For lkey in modifier.keys
      {
        setKeyDown(lkey, TB_Modifier)
      }

      ret_val := 1
    }

    ; Process layers
    For layer_name in layers
    {
      layer := TB_Layers[layer_name]
      complete := 1
      lsize := 0
      ; MsgBox % layer_name " has layer " layer
      For lkey in layer
      {
        lsize := lsize + 1
        ; MsgBox % layer_name " has key " lkey
        if (!GetModKeyState(lkey, "P"))
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

      ; Mark the keys of the layer as down.
      layer := TB_Layers[tmp_layer]
      For lkey in layer
      {
        setKeyDown(lkey, TB_Layer)
      }

      ret_val := 1
    }
  }

  return ret_val
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

; Get key mappings function.
getKey(key) {
  global keySet
  keys := keySet[key]
  return keys
}

getLayerFromAlias(layer_alias)
{
  Global TB_LayerAliases

  tmp_layer := TB_LayerAliases[layer_alias]
  if (StrLen(tmp_layer) <= 0)
  {
    tmp_layer := layer_alias
  }
  
  return tmp_layer
}

getDownKey(key) {
  Global TB_ModKeysDown
  Global TB_ModKeysUp
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

  layer := getLayerFromAlias(layer)

  key_down := layer " Down"

  the_key := keys[key_down]
  If (StrLen(TB_ModKeysDown) > 0)
  {
    the_key := TB_ModKeysDown the_key TB_ModKeysUp
    ; MsgBox % "Will send: " the_key
  }
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
  If (StrLen(TB_ModKeysUp) > 0)
  {
    the_key := the_key TB_ModKeysUp
  }
  return the_key
}

DoHotKeyDown:
  Critical, 1000
  Global TB_KeyDelay
  Global TB_Layer
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

    if (!isComplexComplete(hkey))
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

  return


DoComplexHotKeyDown:
  Critical, 1000
  Global TB_KeyDelay
  Global TB_Layer
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
    if (!isComplexComplete(hkey))
    {
      setKeyDown(hkey, TB_Layer)
    }
  }
  else
  {
    ; MsgBox % "THERE"
  }
  if (!isActiveModifierKey(hkey))
  {
    key := getDownKey(hkey)
  }

  SendInput, %key%
  return

DoComplexHotKeyUp:
  Critical, 1000
  Global TB_LayerKeys
  Global TB_ActiveLayers
  Global TB_Layer
  Global TB_ModKeysDown
  Global TB_ModKeysUp
  Global TB_ActiveModifiers
  Global TB_Modifier
  Global TB_Modifiers
  Global TB_ModifierKeys
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

  ; Remove active modifers for key.
  if (isModifierKey(this_key))
  {
    ; Delete active modifiers for this key.
    modifiers := TB_ModifierKeys[this_key]
    For modifier_name in modifiers
    {
      if (TB_ActiveModifiers[modifier_name] > 0)
      {
        TB_ActiveModifiers.Delete(modifier_name)
        up_key := TB_Modifiers[modifier_name].key_up
      }
    }

    ; Find which modifier is now active.
    active_modifier := ""
    lkey_size := 0
    TB_ModKeysDown := ""
    TB_ModKeysUp := ""
    For modifier in TB_ActiveModifiers
    {
      TB_ModKeysDown := TB_ModKeysDown TB_Modifiers[modifier].key_down
      TB_ModKeysUp := TB_ModKeysUp TB_Modifiers[modifier].key_up
      if (TB_ActiveModifiers[modifier] > lkey_size)
      {
        lkey_size := TB_ActiveModifiers[modifier]
        active_modifier := modifier
      }
    }

    ; Set the active modifier.
    TB_Modifier := active_modifier
  }
  return

DoSpaceHotKeyDown:
  Critical, 1000
  Global TB_KeyDelay
  Global TB_Layer
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
    while (kdelay > 0 AND GetKeyState("Space", "P"))
    {
      ; KeyWait, %hkey%, T%kdelay%
      MySleep(10)
      kdelay := kdelay - 10
    }
    if (GetKeyState("Space", "P"))
    {
      if (!isComplexComplete(hkey))
      {
        setKeyDown(hkey, TB_Layer)
      }
    }
    else
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

DoSpaceHotKeyUp:
  Critical, 1000
  Global TB_LayerKeys
  Global TB_ActiveLayers
  Global TB_Layer
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

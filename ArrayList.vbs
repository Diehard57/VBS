'Thanks for Jeremy England: https://gist.github.com/simply-coded

'create an arraylist
Set list = CreateObject("System.Collections.ArrayList")

'add values
list.add "mike"
list.add "jeremy"
list.add "kate"
list.add "alice"

'add a bunch of values easily via custom sub
addrange list, Split("ben alice mark john alice")

'get an item
MsgBox list(2)
MsgBox list.item(4)

'change an item
list(2) = "katelin"
list.item(4) = "larry"

'count items
MsgBox list.count

'display all values
MsgBox Join(list.toarray)

'sort list
list.sort
list.reverse

'add item if doesn't exist
If Not list.contains("garry") Then
  list.add "garry"
End If

'find index of item
MsgBox list.indexof("alice",0)

'find second appearance of item's index
MsgBox list.indexof("alice",list.indexof("alice",0) + 1)

'insert a new item
list.insert 3, "miller"

'remove an item
list.remove "alice"

'remove item at index
list.removeat 5

'clone the list
Set names = list.clone

'clear the list
list.clear


'Common Methods & Properties
'          .Add "VALUE"
'VALUE   = .Item(INDEX)
'LENGTH  = .Count
'ARRAY   = .ToArray
'          .Sort
'          .Reverse
'BOOLEAN = .Contains("VALUE")
'INDEX   = .IndexOf("VALUE", START_INDEX)
'          .Insert INDEX, "VALUE"
'          .Remove "VALUE"
'          .RemoveAt INDEX
'LIST    = .Clone
'          .Clear

Sub addrange(list, range)
  For Each item In range
    list.add item
  Next
End Sub

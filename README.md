# 🎻 Pugh-Analysis-SCAMPER-Method-in-Manufacturing-using-Excel-VBA
This repository contains realistic manufacturing use cases that apply Pugh analysis and the SCAMPER method for process innovation and improvement. These methods enable teams to generate decision matrices, brainstorm variations and select optimal solutions faster and more reliably

---

## 🎪 Pugh Analysis
- It is a solution to fix a root cause or issue
- It is a decision-making matrix used for comparing and evaluating multiple solution options in relation to a baseline option
- It is used by selecting the most important criteria needed for taking the decision and comparing the alternatives
- It is used when only one solution is possible or when a hybrid of many potential solutions are needed
**Eg**:
  1. The time to implement
  2. The cost to implement
  3. The ease of implementation
  4. Safety or any number of similar things
- If alternative solution is better than the current state, enter "+1"
- If alternative solution is same as the current state, enter "0"
- If alternative solution is worse than the current state, enter "-1"
- In a Solution Prioritizing Matrix, each solution weighted on its own merit to set the criteria and not compared to a baseline or its current state

---

## 🎣 SCAMPER Method
This method helps in finding a solution by asking questions about existing products in a different SCAMPER categories. These questions help us come up with creative ideas for developing new products and for improving current ones

**S** - Substitute

**C** - Combine

**A** - Adapt

**M** - Modify

**P** - Put to Another use

**E** - Eliminate

**R** - Reverse

---

## 🤿 Positive & Negative Brainstorming
- Positive Brainstorming focuses on how to achieve or enhance the goal
- Negative Brainstorming method encourages a team to explore new solutiosn by thinking about things that do not seem to be inherently useful

---
## 👨‍💻 Case study: Enhancing narrow width steel with new annealing-pickling technique
### Context
A steel manufacturer faces compliants from niche customers about durability loss in narrow-width coils post annealing-pickling. The team needs to assess alternative process combinations to balance:
- Durability (meansured in bend test score)
- Surface finish
- Cost
- Cycle time
- Sustainability

### Objective
1. **Pugh Analysis**: To compare multiple process combinations
2. **SCAMPER**: To ideate process improvements beyond conventional choices

### Techniques to evaluate
1. *Conventional Nitrogen Annealing + Acid Pickling*
2. *Hydrogen Annealing + Electro Pickling*
3. *Vacuum Annealing + Citric Acid Bath*
4. *Bright annealing + Low-temperature Pickling*

### Evaluation Criteria (for Pugh matrix)
- Durability
- Surface finish quality
- Cost efficiency
- Cycle time
- Sustainability

### Excel VBA code: Pugh Analysis for Technique selection
-           Sub GeneratePughMatrix_Annealing()
            Dim criteria As Variant, techniques As Variant, baseAlt As String
            Dim i As Integer, j As Integer
            Dim ws As Worksheet
            Dim sheetName As String
    
            ' Define your data
              criteria = Array("Durability", "Surface Finish", "Cost", "Cycle Time", "Sustainability")
              techniques = Array("Nitrogen + Acid", "Hydrogen + Electro", "Vacuum + Citric", "Bright + LowTemp")
              baseAlt = "Nitrogen + Acid"
              sheetName = "Annealing_PughMatrix"
    
            ' Delete sheet if it already exists to avoid 1004 error
              On Error Resume Next
              Application.DisplayAlerts = False
              Worksheets(sheetName).Delete
              Application.DisplayAlerts = True
              On Error GoTo 0
        
            ' Add new worksheet
              Set ws = ThisWorkbook.Sheets.Add
              ws.Name = "Annealing_PughMatrix"
    
            ' Add headers
              ws.Cells(1, 1).Value = "Techniques"
              For j = 0 To UBound(criteria)
                ws.Cells(1, j + 2).Value = criteria(j)
              Next j
    
            'Score inputs
              For i = 0 To UBound(techniques)
              ws.Cells(i + 2, 1).Value = techniques(i)
                For j = 0 To UBound(criteria)
            Dim score As Variant
            score = InputBox("Score " & techniques(i) & " on " & criteria(j) & vbCtrlf & "(-1 worse, 0 same, +1 better)", "Pugh Analysis")
        
        'Validate numeric input
        If IsNumeric(score) Then
            ws.Cells(i + 2, j + 2).Value = score
        Else
            ws.Cells(i + 2, j + 2).Value = 0 ' Default to 0 if user cancels or gives invalid output
        End If
      Next j
      Next i
    
      MsgBox "Pugh Matrix created on sheet: " & ws.Name
   
      End Sub

### Excel VBA code: SCAMPER Idea
-       Sub SCAMPER_AnnealingPickling()
        Dim actions As Variant
        actions = Array("Substitute", "Combine", "Adapt", "Modify", "Put to Another Use", "Eliminate", "Rearrange")
    
        Dim i As Integer, response As String
        Sheets.Add.Name = "SCAMPER Annealing"
        Cells(1, 1).Value = "SCAMPER Action"
        Cells(1, 2).Value = "Idea"
    
        For i = 0 To UBound(actions)
        response = InputBox("Enter idea for: " & actions(i) & vbCtrlf & _
            "Based on improving narrow-width processing in annealing-pickling line", "SCAMPER Prompt")
        Cells(i + 2, 1).Value = actions(i)
        Cells(i + 2, 2).Value = response
      Next i
    
      MsgBox "SCAMPER ideas recorded in sheet: SCAMPER_Annealing"
            
      End Sub

---

## 🧑‍🔬 Excel functionalities used
- VBA macros

---

## 🧑‍⚖️ Requirements
- Microsoft Excel 2016 or later
- Lean & Six Sigma fundamentals

---
*"Excel is a canvas on which data paints its story"*

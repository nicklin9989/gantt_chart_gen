import React, { useState, useRef } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
// 工具函數：產生日期+小時陣列
function getDateHourArray(startDate, endDate, hoursPerDay = 24) {
  const arr = [];
  let d = new Date(startDate);
  while (d <= endDate) {
    for (let h = 0; h < hoursPerDay; h++) {
      arr.push({
        date: d.toISOString().slice(0, 10),
        hour: h,
        label: `${d.toISOString().slice(0, 10)} ${h}:00`,
      });
    }
    d = new Date(d);
    d.setDate(d.getDate() + 1);
  }
  return arr;
}
// 預設任務名稱列表
const defaultTaskNames = [
  "Grid IPRO",
  "IMAP",
  "Ref. scans",
  "Ramping down from XO2",
  "PSU replacement",
  "Vent the column",
  "pinka-11 upgrade",
  "e-source replacement",
  "gun stabilization",
  "Ramp to XO2",
  "XO2 alignment",
  "Burn in, Exposure alignment",
  "Grid IPRO, IMAP",
  "1st Monitor plate",
  "2nd Monitor plate",
  "SCANS + RAMP DOWN",
  "Disassemble housing",
  "LIFT TOOL BASE",
  "Open tool base",
  "Toolbase cleaning",
  "Leak check",
  "Stage test (overnight)",
  "Particle check",
  "Monitor laser ratio",
  "Close tool base",
  "Pump down toolbase",
  "Leak check",
  "Lower column",
  "assemble housing",
  "Vent column",
  "Leak test column",
  "Replace IP33",
  "Partial housing reassembly",
  "Pump down column overnight",
  "Leak test column",
  "Vent interface, Remove shielding",
  "e-source replacement",
  "Re-assmemble housing",
  "Ramp up to XO2",
  "Requal",
  "Software upgrade",
];
// 預設模板
const defaultTemplates = [
  {
    name: "e-source replacement",
    tasks: [
      {
        name: "Grid IPRO, IMAP, Ref. scans",
        start: 0,
        duration: 2,
        color: "#FFB347",
      },
      {
        name: "Ramping down from XO2",
        start: 2,
        duration: 2,
        color: "#77DD77",
      },
      {
        name: "Vent the column, diss housing",
        start: 4,
        duration: 3,
        color: "#AEC6CF",
      },
      { name: "e-source replacement", start: 7, duration: 3, color: "#CFCFC4" },
      { name: "Reassmemble housing", start: 10, duration: 2, color: "#CFCFC4" },
      { name: "gun stabilisation", start: 12, duration: 8, color: "#FF6961" },
      { name: "Ramp to XO2", start: 20, duration: 3, color: "#B39EB5" },
      { name: "XO2 alignment", start: 23, duration: 8, color: "#B39EB5" },
      {
        name: "Burn in, Exposure alignment",
        start: 31,
        duration: 12,
        color: "#B39EB5",
      },
      { name: "Grid IPRO, IMAP", start: 43, duration: 6, color: "#AEC6CF" },
      { name: "1st Monitor plate", start: 49, duration: 5, color: "#CFCFC4" },
    ],
  },
];
// 調色盤
const palette = [
  "#FFB347",
  "#77DD77",
  "#AEC6CF",
  "#CFCFC4",
  "#F49AC2",
  "#B39EB5",
  "#FF6961",
  "#CB99C9",
  "#FFD700", // 金色
  "#40E0D0", // 青綠
  "#FFA07A", // 淺橙
  "#8FBC8F", // 深綠
  "#6495ED", // 藍
  "#FF69B4", // 粉紅
  "#CD5C5C", // 棗紅
  "#20B2AA", // 青藍
];
function App() {
  // 獲取當前日期
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  // 狀態
  const [startDate, setStartDate] = useState(today.toISOString().slice(0, 10));
  const [endDate, setEndDate] = useState(tomorrow.toISOString().slice(0, 10));
  const [hoursPerDay, setHoursPerDay] = useState(24);
  const [tasks, setTasks] = useState([]); // 空陣列，沒有預設任務
  const [newTaskName, setNewTaskName] = useState(defaultTaskNames[0]);
  const [customTaskName, setCustomTaskName] = useState("");
  const [newTaskHours, setNewTaskHours] = useState(8);
  const [selectedColor, setSelectedColor] = useState(palette[0]);
  const [selectedTask, setSelectedTask] = useState(null);
  const [shiftHours, setShiftHours] = useState(0);
  const fileInputRef = useRef();
  const [selectedTemplate, setSelectedTemplate] = useState("");
  const [showHoursColumn, setShowHoursColumn] = useState(true);
  // 產生橫軸
  const allDateHours = getDateHourArray(
    new Date(startDate),
    new Date(endDate),
    hoursPerDay
  );
  // 新增任務
  function addTask() {
    const name = newTaskName === "" ? customTaskName : newTaskName;
    if (!name) return;
    setTasks([
      ...tasks,
      {
        name,
        start: 0,
        duration: newTaskHours,
        color: selectedColor,
      },
    ]);
    setCustomTaskName(""); // 新增後清空自訂名稱
  }
  // 拖曳 bar
  function onBarDrag(idx, delta) {
    setTasks((tasks) =>
      tasks.map((t, i) =>
        i === idx
          ? {
              ...t,
              start: Math.max(
                0,
                Math.min(allDateHours.length - t.duration, t.start + delta)
              ),
            }
          : t
      )
    );
  }
  // 拉伸 bar
  function onBarResize(idx, delta, edge) {
    setTasks((tasks) =>
      tasks.map((t, i) => {
        if (i !== idx) return t;
        if (edge === "left") {
          // 左邊拉伸：固定右邊，改變起始點和長度
          const newStart = Math.max(0, t.start + delta);
          const newDuration = t.duration - (newStart - t.start);
          if (newDuration < 1) return t;
          return { ...t, start: newStart, duration: newDuration };
        } else {
          // 右邊拉伸：固定左邊，只改變長度
          const newDuration = Math.max(1, t.duration + delta);
          if (t.start + newDuration > allDateHours.length) return t;
          return { ...t, duration: newDuration };
        }
      })
    );
  }
  // 選擇任務
  function selectTask(idx) {
    setSelectedTask(idx);
    // 同時更新選中的顏色
    setSelectedColor(tasks[idx].color);
  }
  // 選擇顏色
  function selectColor(color) {
    setSelectedColor(color);
    // 如果有選中的任務，同時更新任務顏色
    if (selectedTask !== null) {
      setTasks((tasks) =>
        tasks.map((t, i) => (i === selectedTask ? { ...t, color } : t))
      );
    }
  }
  // 刪除任務
  function deleteTask() {
    if (selectedTask !== null) {
      setTasks((tasks) => tasks.filter((_, i) => i !== selectedTask));
      setSelectedTask(null);
    }
  }
  // 拖曳任務順序
  function onTaskDrag(fromIdx, toIdx) {
    if (fromIdx === toIdx) return;
    setTasks((tasks) => {
      const newTasks = [...tasks];
      const [moved] = newTasks.splice(fromIdx, 1);
      newTasks.splice(toIdx, 0, moved);
      return newTasks;
    });
  }
  // 匯出 Excel
  async function exportExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Gantt");

    // Header Row 1: Action, Hours (conditional), Dates
    const headerRow1Values = ["Action"];
    if (showHoursColumn) {
      headerRow1Values.push("Hours");
    }
    for (let d = 0; d < allDateHours.length; d += hoursPerDay) {
      headerRow1Values.push(allDateHours[d].date);
      for (let h = 1; h < hoursPerDay; h++) {
        headerRow1Values.push(null);
      }
    }
    const excelHeaderRow1 = sheet.addRow(headerRow1Values);

    // Merge date cells in Header Row 1
    let dateMergeStartCol = showHoursColumn ? 3 : 2;
    for (let d = 0; d < allDateHours.length; d += hoursPerDay) {
      sheet.mergeCells(1, dateMergeStartCol, 1, dateMergeStartCol + hoursPerDay - 1);
      dateMergeStartCol += hoursPerDay;
    }

    // Style Header Row 1
    excelHeaderRow1.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD9D9D9" },
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

    // Header Row 2: Placeholders and Time Markers (00:00, 12:00)
    const excelHeaderRow2Values = [];
    excelHeaderRow2Values.push(null); // Cell under "Action"
    if (showHoursColumn) {
      excelHeaderRow2Values.push(null); // Cell under "Hours"
    }
    allDateHours.forEach(dh => {
      const hourInDay = dh.hour;
      if (hourInDay === 0) {
        excelHeaderRow2Values.push("00:00");
      } else if (hourInDay === 12) {
        excelHeaderRow2Values.push("12:00");
      } else {
        excelHeaderRow2Values.push(null);
      }
    });
    const excelHeaderRow2 = sheet.addRow(excelHeaderRow2Values);

    // Style Header Row 2
    for (let i = 1; i <= sheet.columnCount; i++) {
      const cell = excelHeaderRow2.getCell(i);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF2F2F2" }, // Light grey background
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      if (cell.value === "00:00" || cell.value === "12:00") {
        cell.font = { size: 8, color: { argb: "FF595959" } }; // Smaller, grey font
      }
    }
    
    // Data rows
    tasks.forEach((task) => {
      const rowData = [task.name];
      if (showHoursColumn) {
        rowData.push(task.duration);
      }
      for (let i = 0; i < allDateHours.length; i++) {
        rowData.push(i >= task.start && i < task.start + task.duration ? " " : "");
      }
      const r = sheet.addRow(rowData);

      // Cell coloring for tasks
      const taskCellsStartCol = showHoursColumn ? 3 : 2;
      for (let i = 0; i < allDateHours.length; i++) {
        const cell = r.getCell(i + taskCellsStartCol);
        if (i >= task.start && i < task.start + task.duration) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: task.color.replace("#", "") },
          };
        }
        // Apply thin borders to all data cells in the timeline part
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      }
       // Style Name and Hours cells for data rows
      r.getCell(1).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      if (showHoursColumn) {
        r.getCell(2).alignment = { vertical: "middle", horizontal: "center" };
        r.getCell(2).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      }
    });

    // Column Widths
    sheet.getColumn(1).width = 30; // Action column
    let currentColumnIndex = 2;
    if (showHoursColumn) {
      sheet.getColumn(currentColumnIndex).width = 8; // Hours column
      currentColumnIndex++;
    }
    for (let i = 0; i < allDateHours.length; i++) {
      sheet.getColumn(currentColumnIndex + i).width = 10; // Each hour cell in timeline (調寬)
    }

    // Overall thick outer border (applied after all rows and cells have their individual thin borders)
    const firstRowIdx = 1;
    const lastRowIdx = sheet.rowCount;
    const firstColIdx = 1;
    const lastColIdx = sheet.columnCount;

    for (let r = firstRowIdx; r <= lastRowIdx; r++) {
      for (let c = firstColIdx; c <= lastColIdx; c++) {
        const cell = sheet.getRow(r).getCell(c);
        // Ensure all cells have a base thin border object to modify
        if (!cell.border) {
            cell.border = {};
        }
        const currentBorder = JSON.parse(JSON.stringify(cell.border)); // Deep copy

        currentBorder.top = r === firstRowIdx ? { style: "thick" } : (currentBorder.top || { style: "thin" });
        currentBorder.bottom = r === lastRowIdx ? { style: "thick" } : (currentBorder.bottom || { style: "thin" });
        currentBorder.left = c === firstColIdx ? { style: "thick" } : (currentBorder.left || { style: "thin" });
        currentBorder.right = c === lastColIdx ? { style: "thick" } : (currentBorder.right || { style: "thin" });
        cell.border = currentBorder;
      }
    }
    
    // Download
    const buf = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buf]), "gantt.xlsx");
  }
  // 匯入 Excel
  async function importExcel(e) {
    const file = e.target.files[0];
    if (!file) return;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheet = workbook.worksheets[0];
    const rows = [];
    sheet.eachRow((row) => rows.push(row.values));
    // 解析
    const header = rows[0];
    const hourRow = rows[1];
    const dataRows = rows.slice(2);
    // 重新對齊
    const newTasks = dataRows.map((row) => {
      const name = row[1];
      let start = null;
      let duration = 0;
      for (let i = 2; i < row.length; i++) {
        if (row[i] !== "" && start === null) start = i - 2;
        if (row[i] !== "") duration++;
      }
      // 顏色
      let color = "#FFB347";
      const cell = sheet.getRow(rows.indexOf(row) + 1).getCell(start + 2);
      if (cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb) {
        color = "#" + cell.fill.fgColor.argb;
      }
      return { name, start, duration, color };
    });
    setTasks(newTasks);
    setSelectedTask(null);
  }
  // 鍵盤刪除
  React.useEffect(() => {
    function onKeyDown(e) {
      if (e.key === "Delete") deleteTask();
    }
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });
  // 拖曳 bar & resize
  function handleBarMouseDown(idx, e, edge) {
    e.preventDefault();
    let startX = e.clientX;
    let lastDelta = 0;
    function onMove(ev) {
      const moveX = ev.clientX - startX;
      const delta = Math.floor(moveX / 40);
      if (delta !== lastDelta) {
        if (edge) {
          onBarResize(idx, delta - lastDelta, edge);
        } else {
          onBarDrag(idx, delta - lastDelta);
        }
        lastDelta = delta;
      }
    }
    function onUp() {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    }
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }
  // 拖曳任務順序
  function handleTaskDragStart(idx, e) {
    e.dataTransfer.setData("taskIdx", idx);
  }
  function handleTaskDrop(idx, e) {
    const fromIdx = Number(e.dataTransfer.getData("taskIdx"));
    onTaskDrag(fromIdx, idx);
  }
  // 平移所有任務
  function shiftAllTasks(delta) {
    setTasks((tasks) =>
      tasks.map((t) => {
        let newStart = t.start + delta;
        // 限制不能小於 0，也不能超過橫軸最大格
        newStart = Math.max(
          0,
          Math.min(allDateHours.length - t.duration, newStart)
        );
        return { ...t, start: newStart };
      })
    );
  }
  // 應用模板
  function applyTemplate(templateName) {
    const template = defaultTemplates.find((t) => t.name === templateName);
    if (template) {
      setTasks(template.tasks);
      setSelectedTemplate(templateName);
    }
  }
  // UI
  return (
    <div style={{ padding: 24, fontFamily: "sans-serif" }}>
      <h2>IMS gantt chart generator</h2>
      <p style={{ fontSize: 14, color: "#666", margin: "4px 0 16px" }}>
        Creator: Nick Lin
      </p>
      <div style={{ marginBottom: 12 }}>
        <label>
          choose template:
          <select
            value={selectedTemplate}
            onChange={(e) => applyTemplate(e.target.value)}
            style={{ margin: "0 8px" }}
          >
            <option value="">-- choose template --</option>
            {defaultTemplates.map((template) => (
              <option key={template.name} value={template.name}>
                {template.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          Project Start:
          <input
            type="date"
            value={startDate}
            onChange={(e) => setStartDate(e.target.value)}
            style={{ margin: "0 8px" }}
          />
        </label>
        <label>
          Project End:
          <input
            type="date"
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
            style={{ margin: "0 8px" }}
          />
        </label>
        <label>
          Cells/Day:
          <input
            type="number"
            min={1}
            max={24}
            value={hoursPerDay}
            onChange={(e) => setHoursPerDay(Number(e.target.value))}
            style={{ width: 40, margin: "0 8px" }}
          />
        </label>
        <button onClick={exportExcel}>Export Excel</button>
        <input
          type="file"
          accept=".xlsx"
          ref={fileInputRef}
          style={{ display: "none" }}
          onChange={importExcel}
        />
        <button onClick={() => fileInputRef.current.click()}>
          Import Excel
        </button>
        <label style={{ marginLeft: "8px" }}>
          <input
            type="checkbox"
            checked={showHoursColumn}
            onChange={(e) => setShowHoursColumn(e.target.checked)}
          />
          Show Hours Column
        </label>
      </div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
        <span>Palette:</span>
        {palette.map((c) => (
          <div
            key={c}
            onClick={() => selectColor(c)}
            style={{
              width: 24,
              height: 24,
              background: c,
              border: c === selectedColor ? "2px solid #333" : "1px solid #ccc",
              margin: "0 4px",
              cursor: "pointer",
            }}
          />
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
        <select
          value={newTaskName}
          onChange={(e) => setNewTaskName(e.target.value)}
        >
          {defaultTaskNames.map((n) => (
            <option key={n} value={n}>
              {n}
            </option>
          ))}
          <option value="">Custom...</option>
        </select>
        {newTaskName === "" && (
          <input
            placeholder="Task name"
            value={customTaskName}
            onChange={(e) => setCustomTaskName(e.target.value)}
            style={{ marginLeft: 8 }}
          />
        )}
        <input
          type="number"
          min={1}
          max={allDateHours.length}
          value={newTaskHours}
          onChange={(e) => setNewTaskHours(Number(e.target.value))}
          style={{ width: 40, margin: "0 8px" }}
        />
        <button onClick={addTask}>Add Task</button>
      </div>
      <div style={{ margin: "12px 0" }}>
        <button onClick={() => shiftAllTasks(-1)}>←</button>
        <span style={{ margin: "0 8px" }}>Shift All Tasks</span>
        <button onClick={() => shiftAllTasks(1)}>→</button>
      </div>
      {/* Gantt Chart Table */}
      <div
        style={{
          overflowX: "auto",
          border: "1px solid #ccc",
          width: "100%",
        }}
      >
        <table
          style={{
            borderCollapse: "collapse",
            minWidth: 800,
            tableLayout: "fixed",
          }}
        >
          <thead>
            <tr>
              <th
                style={{
                  width: 200,
                  background: "#eee",
                  position: "sticky",
                  left: 0,
                  zIndex: 1,
                  border: "1px solid #ccc",
                }}
              >
                Action
              </th>
              {showHoursColumn && (
                <th
                  style={{
                    width: 60,
                    background: "#eee",
                    textAlign: "center",
                    border: "1px solid #ccc",
                  }}
                >
                  Hours
                </th>
              )}
              {allDateHours.map((dh, i) =>
                i % hoursPerDay === 0 ? (
                  <th
                    key={i}
                    colSpan={hoursPerDay}
                    style={{
                      background: "#f9f9f9",
                      border: "1px solid #ccc",
                      textAlign: "center",
                      width: hoursPerDay * 40,
                      minWidth: hoursPerDay * 40,
                      maxWidth: hoursPerDay * 40,
                    }}
                  >
                    {dh.date}
                  </th>
                ) : null
              )}
            </tr>
            {/* Modified Second Header Row for Time Markers */}
            <tr>
              <th
                style={{
                  width: 200,
                  background: "#f0f0f0",
                  position: "sticky",
                  left: 0,
                  zIndex: 1,
                  border: "1px solid #ccc",
                }}
              />
              {showHoursColumn && (
                <th
                  style={{
                    width: 60,
                    background: "#f0f0f0",
                    textAlign: "center",
                    border: "1px solid #ccc",
                  }}
                />
              )}
              {allDateHours.map((dh, index) => {
                const hourInDay = dh.hour;
                let label = "";
                if (hourInDay === 0) {
                  label = "00:00";
                } else if (hourInDay === 12) {
                  label = "12:00";
                }
                return (
                  <th
                    key={`header-hour-marker-${index}`}
                    style={{
                      background: "#f0f0f0",
                      border: "1px solid #ccc",
                      textAlign: label === "00:00" ? "left" : "center",
                      paddingLeft: label === "00:00" ? "2px" : undefined,
                      width: 40,
                      minWidth: 40,
                      maxWidth: 40,
                      fontSize: "10px",
                      fontWeight: "normal",
                    }}
                  >
                    {label}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {tasks.map((task, idx) => (
              <tr
                key={idx}
                draggable
                onDragStart={(e) => handleTaskDragStart(idx, e)}
                onDrop={(e) => handleTaskDrop(idx, e)}
                onDragOver={(e) => e.preventDefault()}
                style={{
                  background: selectedTask === idx ? "#e0e0ff" : "white",
                  cursor: "pointer",
                }}
                onClick={() => selectTask(idx)}
              >
                <td
                  style={{
                    border: "1px solid #ccc",
                    width: 200,
                    minWidth: 200,
                    maxWidth: 200,
                    position: "sticky",
                    left: 0,
                    background: selectedTask === idx ? "#e0e0ff" : "white",
                    zIndex: 1,
                  }}
                >
                  {task.name}
                </td>
                {showHoursColumn && (
                  <td
                    style={{
                      border: "1px solid #ccc",
                      width: 60,
                      textAlign: "center",
                    }}
                  >
                    {task.duration}
                  </td>
                )}
                {allDateHours.map((dh, i) => {
                  // bar
                  if (i === task.start) {
                    return (
                      <td
                        key={i}
                        colSpan={task.duration}
                        style={{
                          background: task.color,
                          border: "2px solid #333",
                          position: "relative",
                          cursor: "move",
                          width: task.duration * 40,
                          minWidth: task.duration * 40,
                          maxWidth: task.duration * 40,
                        }}
                        onMouseDown={(e) => handleBarMouseDown(idx, e)}
                      >
                        {/* 左右拉伸手柄 */}
                        <span
                          style={{
                            position: "absolute",
                            left: 0,
                            top: 0,
                            width: 12,
                            height: "100%",
                            cursor: "ew-resize",
                            background: "rgba(0,0,0,0.1)",
                          }}
                          onMouseDown={(e) =>
                            handleBarMouseDown(idx, e, "left")
                          }
                        />
                        <span
                          style={{
                            position: "absolute",
                            right: 0,
                            top: 0,
                            width: 12,
                            height: "100%",
                            cursor: "ew-resize",
                            background: "rgba(0,0,0,0.1)",
                          }}
                          onMouseDown={(e) =>
                            handleBarMouseDown(idx, e, "right")
                          }
                        />
                      </td>
                    );
                  }
                  // bar 內部已合併
                  if (i > task.start && i < task.start + task.duration)
                    return null;
                  // 其他空格
                  return (
                    <td
                      key={i}
                      style={{
                        border: "1px solid #ccc",
                        width: 40,
                        minWidth: 40,
                        maxWidth: 40,
                      }}
                    />
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: 8, color: "#888" }}>
        <ul>
          <li>
            Drag the bar to change the start time; drag the edges to adjust the
            duration
          </li>
          <li>Select a task and press Delete to remove it</li>
          <li>Drag a task row to rearrange the order</li>
          <li>Select a task and click a color to change its color</li>
        </ul>
      </div>
    </div>
  );
}

export default App;

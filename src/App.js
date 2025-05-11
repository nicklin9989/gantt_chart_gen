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
  const [newTaskHours, setNewTaskHours] = useState(8);
  const [selectedColor, setSelectedColor] = useState(palette[0]);
  const [selectedTask, setSelectedTask] = useState(null);

  const fileInputRef = useRef();

  // 產生橫軸
  const allDateHours = getDateHourArray(
    new Date(startDate),
    new Date(endDate),
    hoursPerDay
  );

  // 新增任務
  function addTask() {
    setTasks([
      ...tasks,
      {
        name: newTaskName,
        start: 0,
        duration: newTaskHours,
        color: selectedColor,
      },
    ]);
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

    // 標題列
    const headerRow1 = ["Action"];
    let lastDate = "";
    allDateHours.forEach((dh, i) => {
      if (dh.date !== lastDate) {
        headerRow1.push(`${dh.date}`);
        lastDate = dh.date;
      } else {
        headerRow1.push(null);
      }
    });
    sheet.addRow(headerRow1);

    // 小時列
    const headerRow2 = [" "];
    allDateHours.forEach((dh, i) => {
      headerRow2.push(i % hoursPerDay === 0 ? hoursPerDay : null);
    });
    sheet.addRow(headerRow2);

    // 合併日期儲存格
    let col = 2;
    for (let d = 0; d < allDateHours.length; d += hoursPerDay) {
      sheet.mergeCells(1, col, 1, col + hoursPerDay - 1);
      col += hoursPerDay;
    }

    // 資料列
    tasks.forEach((task) => {
      const row = [task.name];
      for (let i = 0; i < allDateHours.length; i++) {
        row.push(i >= task.start && i < task.start + task.duration ? " " : "");
      }
      const r = sheet.addRow(row);
      // 色塊樣式
      for (let i = 0; i < allDateHours.length; i++) {
        const cell = r.getCell(i + 2);
        if (i >= task.start && i < task.start + task.duration) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: task.color.replace("#", "") },
          };
        }
      }
    });

    // 欄寬
    sheet.columns.forEach((col, i) => {
      col.width = i === 0 ? 30 : 8;
    });

    // 下載
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

  // UI
  return (
    <div style={{ padding: 24, fontFamily: "sans-serif" }}>
      <h2>IMS gantt chart generator</h2>
      <p style={{ fontSize: 14, color: "#666", margin: "4px 0 16px" }}>
    Creator: Nick Lin
  </p>
      <div style={{ marginBottom: 12 }}>
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
          Hours/Day:
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
            value={newTaskName}
            onChange={(e) => setNewTaskName(e.target.value)}
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
                }}
              >
                Action
              </th>
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
            <tr>
              <th
                style={{
                  position: "sticky",
                  left: 0,
                  zIndex: 1,
                }}
              />
              {allDateHours.map((dh, i) =>
                i % hoursPerDay === 0 ? (
                  <th
                    key={i}
                    colSpan={hoursPerDay}
                    style={{
                      background: "#f0f0f0",
                      border: "1px solid #ccc",
                      textAlign: "center",
                      width: hoursPerDay * 40,
                      minWidth: hoursPerDay * 40,
                      maxWidth: hoursPerDay * 40,
                    }}
                  >
                    {hoursPerDay}
                  </th>
                ) : null
              )}
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

<<<<<<< HEAD
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
=======
import React, { useState } from "react";
import * as ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { DragDropContext, Droppable, Draggable } from "react-beautiful-dnd";

// Default task options
const defaultTaskNames = [
  "Grid IPRO, IMAP, Ref. scans",
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

const colorOptions = [
  "#4f8cff",
  "#34c759",
  "#ffcc00",
  "#ff9500",
  "#ff3b30",
  "#af52de",
  "#5ac8fa",
  "#5856d6",
];

// 產生日期+小時陣列
function getDateHourArray(start, end, hoursPerDay) {
  const arr = [];
  let dt = new Date(start);
  while (dt <= end) {
    for (let h = 0; h < hoursPerDay; h++) {
      arr.push({
        date: dt.toISOString().slice(0, 10),
        hour: h,
        dayLabel: dt.toLocaleDateString("en-US", {
          month: "short",
          day: "numeric",
        }),
        weekLabel: dt.toLocaleDateString("en-US", { weekday: "short" }),
      });
    }
    dt.setDate(dt.getDate() + 1);
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62
  }
  return arr;
}

<<<<<<< HEAD
// 預設任務名稱列表
const defaultTaskNames = [
  "Grid IPRO", "IMAP", "Ref. scans", "Ramping down from XO2", "PSU replacement",
  "Vent the column", "pinka-11 upgrade", "e-source replacement", "gun stabilization",
  "Ramp to XO2", "XO2 alignment", "Burn in, Exposure alignment", "Grid IPRO, IMAP",
  "1st Monitor plate", "2nd Monitor plate", "SCANS + RAMP DOWN", "Disassemble housing",
  "LIFT TOOL BASE", "Open tool base", "Toolbase cleaning", "Leak check", "Stage test (overnight)",
  "Particle check", "Monitor laser ratio", "Close tool base", "Pump down toolbase", "Leak check",
  "Lower column", "assemble housing", "Vent column", "Leak test column", "Replace IP33",
  "Partial housing reassembly", "Pump down column overnight", "Leak test column",
  "Vent interface, Remove shielding", "e-source replacement", "Re-assmemble housing",
  "Ramp up to XO2", "Requal", "Software upgrade"
];

// 調色盤
const palette = [
  "#FFB347", "#77DD77", "#AEC6CF", "#CFCFC4", "#F49AC2", "#B39EB5", "#FF6961", "#CB99C9"
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
  const allDateHours = getDateHourArray(new Date(startDate), new Date(endDate), hoursPerDay);

  // 新增任務
  function addTask() {
    setTasks([
      ...tasks,
      {
        name: newTaskName,
        start: 0,
        duration: newTaskHours,
        color: selectedColor
      }
    ]);
  }

  // 拖曳 bar
  function onBarDrag(idx, delta) {
    setTasks(tasks =>
      tasks.map((t, i) =>
        i === idx
          ? {
              ...t,
              start: Math.max(0, Math.min(allDateHours.length - t.duration, t.start + delta))
=======
// 產生日期分組資訊
function getDateGroups(start, end, hoursPerDay) {
  const arr = [];
  let dt = new Date(start);
  while (dt <= end) {
    arr.push({
      date: dt.toISOString().slice(0, 10),
      dayLabel: dt.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
      }),
      weekLabel: dt.toLocaleDateString("en-US", { weekday: "short" }),
    });
    dt.setDate(dt.getDate() + 1);
  }
  return arr;
}

// 匯出 Excel（有顏色有排版）
async function exportExcelWithStyle(
  tasks,
  allDateHours,
  dateGroups,
  hoursPerDay
) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Gantt");

  // 標題列
  const headerRow1 = ["Action"];
  dateGroups.forEach((dg) => {
    headerRow1.push(`${dg.dayLabel} (${dg.weekLabel})`);
    for (let i = 1; i < hoursPerDay; i++) headerRow1.push(null);
  });
  sheet.addRow(headerRow1);

  // 小時列
  const headerRow2 = [" "].concat(allDateHours.map((dh) => dh.hour));
  sheet.addRow(headerRow2);

  // 合併日期儲存格
  let col = 2;
  dateGroups.forEach(() => {
    sheet.mergeCells(1, col, 1, col + hoursPerDay - 1);
    col += hoursPerDay;
  });

  // 標題樣式
  sheet.getRow(1).eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCC" },
    };
    cell.font = { bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  sheet.getRow(2).eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "F0F8FF" },
    };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });

  // 資料列
  tasks.forEach((task) => {
    const row = [task.name];
    allDateHours.forEach((dh, i) => {
      const inBar = i >= task.start && i < task.start + task.duration;
      row.push(inBar ? " " : "");
    });
    const r = sheet.addRow(row);
    // 標題欄樣式
    r.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFEE" },
    };
    r.getCell(1).alignment = { vertical: "middle", horizontal: "left" };
    r.getCell(1).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
    // 色塊樣式
    allDateHours.forEach((dh, i) => {
      const cell = r.getCell(i + 2);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if (i >= task.start && i < task.start + task.duration) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: task.color.replace("#", "") },
        };
      }
    });
  });

  // 欄寬
  sheet.columns.forEach((col, i) => {
    col.width = i === 0 ? 30 : 5;
  });

  // 下載
  const buf = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buf]), "gantt.xlsx");
}

// 匯入 Excel
function parseGanttExcel(json) {
  const header = json[0];
  const tasks = [];
  for (let i = 2; i < json.length; i++) {
    const row = json[i];
    const name = row[0];
    if (!name) continue;
    let start = null,
      duration = 0;
    let color = "#4f8cff"; // 預設色
    for (let j = 1; j < row.length; j++) {
      if (row[j]) {
        if (start === null) start = j - 1;
        duration++;
      } else if (start !== null && duration > 0) {
        break;
      }
    }
    if (start !== null && duration > 0) {
      tasks.push({
        id: Date.now() + Math.random(),
        name,
        start,
        duration,
        color,
      });
    }
  }
  return tasks;
}

export default function App() {
  const [projectStart, setProjectStart] = useState("");
  const [projectEnd, setProjectEnd] = useState("");
  const [hoursPerDay, setHoursPerDay] = useState(8);

  const [tasks, setTasks] = useState([]);
  const [selectedDefault, setSelectedDefault] = useState("");
  const [customName, setCustomName] = useState("");
  const [duration, setDuration] = useState(2);
  const [selectedColor, setSelectedColor] = useState(colorOptions[0]);
  const [selectedTaskId, setSelectedTaskId] = useState(null);

  // 橫軸
  const allDateHours =
    projectStart && projectEnd
      ? getDateHourArray(
          new Date(projectStart),
          new Date(projectEnd),
          hoursPerDay
        )
      : [];
  const dateGroups =
    projectStart && projectEnd
      ? getDateGroups(new Date(projectStart), new Date(projectEnd), hoursPerDay)
      : [];

  // 新增任務
  const addTask = () => {
    const name = selectedDefault ? selectedDefault : customName;
    if (!name || duration <= 0) return;
    setTasks([
      ...tasks,
      {
        id: Date.now() + Math.random(),
        name,
        start: 0,
        duration,
        color: selectedColor,
      },
    ]);
    setCustomName("");
    setSelectedDefault("");
    setDuration(2);
  };

  // bar 拖曳（移動到任意小時格）
  const onBarMove = (id, newStart) => {
    setTasks((tasks) =>
      tasks.map((t) =>
        t.id === id
          ? {
              ...t,
              start: Math.max(
                0,
                Math.min(newStart, allDateHours.length - t.duration)
              ),
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62
            }
          : t
      )
    );
<<<<<<< HEAD
  }

  // 拉伸 bar
  function onBarResize(idx, delta, edge) {
    setTasks(tasks =>
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
      setTasks(tasks =>
        tasks.map((t, i) =>
          i === selectedTask
            ? { ...t, color }
            : t
        )
      );
    }
  }

  // 刪除任務
  function deleteTask() {
    if (selectedTask !== null) {
      setTasks(tasks => tasks.filter((_, i) => i !== selectedTask));
      setSelectedTask(null);
    }
  }

  // 拖曳任務順序
  function onTaskDrag(fromIdx, toIdx) {
    if (fromIdx === toIdx) return;
    setTasks(tasks => {
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
            fgColor: { argb: task.color.replace("#", "") }
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
=======
  };

  // 只允許拉伸
  const onBarResize = (id, newDuration, direction) => {
    setTasks((tasks) =>
      tasks.map((t) => {
        if (t.id !== id) return t;
        if (direction === "right") {
          // 右側拉伸
          return {
            ...t,
            duration: Math.max(
              1,
              Math.min(newDuration, allDateHours.length - t.start)
            ),
          };
        } else {
          // 左側拉伸
          const diff = t.duration - newDuration;
          const newStart = Math.max(0, t.start + diff);
          return {
            ...t,
            start: newStart,
            duration: Math.max(
              1,
              Math.min(newDuration, allDateHours.length - newStart)
            ),
          };
        }
      })
    );
  };

  // 刪除任務
  const deleteTask = (id) => {
    setTasks((tasks) => tasks.filter((t) => t.id !== id));
    setSelectedTaskId(null);
  };

  // 匯入 Excel
  function handleImportExcel(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      const data = evt.target.result;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(data);
      const worksheet = workbook.worksheets[0];
      const json = [];
      worksheet.eachRow((row, rowNumber) => {
        json.push(row.values.slice(1));
      });
      const importedTasks = parseGanttExcel(json);
      setTasks(importedTasks);
      // 不要自動 setProjectStart、setProjectEnd、setHoursPerDay
    };
    reader.readAsArrayBuffer(file);
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62
  }

  // 鍵盤刪除
  React.useEffect(() => {
<<<<<<< HEAD
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
      <div style={{ marginBottom: 12 }}>
        <label>
          Project Start:
          <input
            type="date"
            value={startDate}
            onChange={e => setStartDate(e.target.value)}
            style={{ margin: "0 8px" }}
          />
        </label>
        <label>
          Project End:
          <input
            type="date"
            value={endDate}
            onChange={e => setEndDate(e.target.value)}
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
            onChange={e => setHoursPerDay(Number(e.target.value))}
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
        <button onClick={() => fileInputRef.current.click()}>Import Excel</button>
      </div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
        <span>Palette:</span>
        {palette.map(c => (
          <div
            key={c}
            onClick={() => selectColor(c)}
            style={{
              width: 24,
              height: 24,
              background: c,
              border: c === selectedColor ? "2px solid #333" : "1px solid #ccc",
              margin: "0 4px",
              cursor: "pointer"
            }}
          />
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
        <select
          value={newTaskName}
          onChange={e => setNewTaskName(e.target.value)}
        >
          {defaultTaskNames.map(n => (
            <option key={n} value={n}>{n}</option>
          ))}
          <option value="">Custom...</option>
        </select>
        {newTaskName === "" && (
          <input
            placeholder="Task name"
            value={newTaskName}
            onChange={e => setNewTaskName(e.target.value)}
          />
        )}
=======
    const handleKeyDown = (e) => {
      if ((e.key === "Delete" || e.key === "Backspace") && selectedTaskId) {
        deleteTask(selectedTaskId);
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [selectedTaskId]);

  // 拖曳排序
  const onDragEnd = (result) => {
    if (!result.destination) return;
    const newTasks = Array.from(tasks);
    const [removed] = newTasks.splice(result.source.index, 1);
    newTasks.splice(result.destination.index, 0, removed);
    setTasks(newTasks);
  };

  // 拖曳 bar
  const handleBarDrag = (task, e) => {
    e.preventDefault();
    const startX = e.clientX;
    const startStart = task.start;
    const onMove = (moveEvent) => {
      const diff = Math.round((moveEvent.clientX - startX) / 40);
      onBarMove(task.id, startStart + diff);
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  return (
    <div
      style={{ maxWidth: 1200, margin: "40px auto", fontFamily: "sans-serif" }}
    >
      <h2>IMS gantt chart generator</h2>
      <div style={{ marginBottom: 16 }}>
        <label>Project start: </label>
        <input
          type="date"
          value={projectStart}
          onChange={(e) => setProjectStart(e.target.value)}
        />
        <label style={{ marginLeft: 8 }}>End: </label>
        <input
          type="date"
          value={projectEnd}
          onChange={(e) => setProjectEnd(e.target.value)}
        />
        <label style={{ marginLeft: 8 }}>Hours per day: </label>
        <input
          type="number"
          min={1}
          max={24}
          value={hoursPerDay}
          onChange={(e) => setHoursPerDay(Number(e.target.value))}
          style={{ width: 50 }}
        />
      </div>
      {/* Color palette */}
      <div style={{ marginBottom: 8 }}>
        <span style={{ marginRight: 8 }}>Task color: </span>
        {colorOptions.map((c, i) => (
          <span
            key={i}
            onClick={() => setSelectedColor(c)}
            style={{
              display: "inline-block",
              width: 24,
              height: 24,
              background: c,
              border: selectedColor === c ? "3px solid #000" : "1px solid #888",
              borderRadius: 4,
              cursor: "pointer",
              marginRight: 4,
            }}
            title={c}
          />
        ))}
      </div>
      <div style={{ marginBottom: 16 }}>
        <select
          value={selectedDefault}
          onChange={(e) => {
            setSelectedDefault(e.target.value);
            setCustomName("");
          }}
          style={{ marginRight: 8 }}
        >
          <option value="">Select default task</option>
          {defaultTaskNames.map((name, i) => (
            <option value={name} key={i}>
              {name}
            </option>
          ))}
        </select>
        <input
          type="text"
          placeholder="Custom task name"
          value={customName}
          onChange={(e) => {
            setCustomName(e.target.value);
            setSelectedDefault("");
          }}
          style={{ marginRight: 8 }}
        />
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62
        <input
          type="number"
          min={1}
          max={allDateHours.length}
<<<<<<< HEAD
          value={newTaskHours}
          onChange={e => setNewTaskHours(Number(e.target.value))}
          style={{ width: 40, margin: "0 8px" }}
        />
        <button onClick={addTask}>Add Task</button>
      </div>
      {/* Gantt Chart Table */}
      <div style={{ 
        overflowX: "auto", 
        border: "1px solid #ccc",
        width: "100%"
      }}>
        <table style={{ 
          borderCollapse: "collapse", 
          minWidth: 800,
          tableLayout: "fixed"
        }}>
          <thead>
            <tr>
              <th style={{ 
                width: 200, 
                background: "#eee",
                position: "sticky",
                left: 0,
                zIndex: 1
              }}>Action</th>
              {allDateHours.map((dh, i) =>
                i % hoursPerDay === 0 ? (
=======
          value={duration}
          onChange={(e) => setDuration(Number(e.target.value))}
          style={{ width: 60, marginRight: 8 }}
        />{" "}
        hours
        <button onClick={addTask}>Add task</button>
        <button
          onClick={() =>
            exportExcelWithStyle(tasks, allDateHours, dateGroups, hoursPerDay)
          }
          style={{ marginLeft: 16 }}
        >
          Export Excel
        </button>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleImportExcel}
          style={{ marginLeft: 16 }}
        />
      </div>
      {allDateHours.length > 0 && (
        <div
          style={{
            overflowX: "auto",
            border: "1px solid #ccc",
            background: "#fafafa",
          }}
        >
          <table style={{ borderCollapse: "collapse", minWidth: 800 }}>
            <thead>
              <tr>
                <th
                  style={{
                    background: "#ffffcc",
                    border: "1px solid #000",
                    minWidth: 120,
                    verticalAlign: "bottom",
                  }}
                  rowSpan={2}
                >
                  Action
                </th>
                {dateGroups.map((dg, i) => (
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62
                  <th
                    key={i}
                    colSpan={hoursPerDay}
                    style={{
<<<<<<< HEAD
                      background: "#f9f9f9",
                      border: "1px solid #ccc",
                      textAlign: "center",
                      width: hoursPerDay * 40,
                      minWidth: hoursPerDay * 40,
                      maxWidth: hoursPerDay * 40
                    }}
                  >
                    {dh.date}
                  </th>
                ) : null
              )}
            </tr>
            <tr>
              <th style={{ 
                position: "sticky",
                left: 0,
                zIndex: 1
              }} />
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
                      maxWidth: hoursPerDay * 40
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
                onDragStart={e => handleTaskDragStart(idx, e)}
                onDrop={e => handleTaskDrop(idx, e)}
                onDragOver={e => e.preventDefault()}
                style={{
                  background: selectedTask === idx ? "#e0e0ff" : "white",
                  cursor: "pointer"
                }}
                onClick={() => selectTask(idx)}
              >
                <td style={{ 
                  border: "1px solid #ccc", 
                  width: 200,
                  minWidth: 200,
                  maxWidth: 200,
                  position: "sticky",
                  left: 0,
                  background: selectedTask === idx ? "#e0e0ff" : "white",
                  zIndex: 1
                }}>
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
                          maxWidth: task.duration * 40
                        }}
                        onMouseDown={e => handleBarMouseDown(idx, e)}
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
                            background: "rgba(0,0,0,0.1)"
                          }}
                          onMouseDown={e => handleBarMouseDown(idx, e, "left")}
                        />
                        <span
                          style={{
                            position: "absolute",
                            right: 0,
                            top: 0,
                            width: 12,
                            height: "100%",
                            cursor: "ew-resize",
                            background: "rgba(0,0,0,0.1)"
                          }}
                          onMouseDown={e => handleBarMouseDown(idx, e, "right")}
                        />
                      </td>
                    );
                  }
                  // bar 內部已合併
                  if (i > task.start && i < task.start + task.duration) return null;
                  // 其他空格
                  return (
                    <td 
                      key={i} 
                      style={{ 
                        border: "1px solid #ccc",
                        width: 40,
                        minWidth: 40,
                        maxWidth: 40
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
=======
                      background: "#e6f7ff",
                      border: "1px solid #000",
                      minWidth: 40 * hoursPerDay,
                      fontSize: 12,
                    }}
                  >
                    {dg.dayLabel} <br />
                    {dg.weekLabel}
                  </th>
                ))}
              </tr>
              <tr>
                {allDateHours.map((dh, i) => (
                  <th
                    key={i}
                    style={{
                      background: "#f0f8ff",
                      border: "1px solid #000",
                      minWidth: 40,
                      fontSize: 12,
                    }}
                  >
                    {dh.hour}
                  </th>
                ))}
              </tr>
            </thead>
            <DragDropContext onDragEnd={onDragEnd}>
              <Droppable droppableId="tasks" direction="vertical">
                {(provided) => (
                  <tbody ref={provided.innerRef} {...provided.droppableProps}>
                    {tasks.map((task, idx) => (
                      <Draggable
                        key={task.id}
                        draggableId={String(task.id)}
                        index={idx}
                      >
                        {(provided2) => (
                          <tr
                            ref={provided2.innerRef}
                            {...provided2.draggableProps}
                            style={{
                              ...provided2.draggableProps.style,
                              background:
                                selectedTaskId === task.id
                                  ? "#ffe066"
                                  : undefined,
                            }}
                          >
                            <td
                              style={{
                                border: "1px solid #000",
                                background: "#ffffee",
                                cursor: "grab",
                              }}
                              {...provided2.dragHandleProps}
                              onClick={() => setSelectedTaskId(task.id)}
                            >
                              {task.name}
                            </td>
                            {allDateHours.map((dh, i) => {
                              const inBar =
                                i >= task.start &&
                                i < task.start + task.duration;
                              const isLeftEdge = inBar && i === task.start;
                              const isRightEdge =
                                inBar && i === task.start + task.duration - 1;
                              return (
                                <td
                                  key={i}
                                  style={{
                                    border: "1px solid #000",
                                    background: inBar ? task.color : "#fff",
                                    position: "relative",
                                    padding: 0,
                                    minWidth: 40,
                                    cursor: inBar ? "pointer" : "default",
                                  }}
                                  onClick={() => setSelectedTaskId(task.id)}
                                >
                                  {/* bar 本體可拖曳 */}
                                  {inBar && isLeftEdge && (
                                    <div
                                      style={{
                                        position: "absolute",
                                        left: 8,
                                        top: 2,
                                        width: (task.duration - 1) * 40,
                                        height: "80%",
                                        background: task.color,
                                        borderRadius: 4,
                                        cursor: "move",
                                        zIndex: 2,
                                      }}
                                      title="Drag to move"
                                      onMouseDown={(e) =>
                                        handleBarDrag(task, e)
                                      }
                                    />
                                  )}
                                  {/* 左側拉伸 */}
                                  {isLeftEdge && (
                                    <div
                                      style={{
                                        position: "absolute",
                                        left: 0,
                                        top: 0,
                                        width: 8,
                                        height: "100%",
                                        cursor: "ew-resize",
                                        background: "#333",
                                      }}
                                      title="Resize"
                                      onMouseDown={(e) => {
                                        e.preventDefault();
                                        const startX = e.clientX;
                                        const startDuration = task.duration;
                                        const startStart = task.start;
                                        const onMove = (moveEvent) => {
                                          const diff = Math.round(
                                            (startX - moveEvent.clientX) / 40
                                          );
                                          onBarResize(
                                            task.id,
                                            startDuration + diff,
                                            "left"
                                          );
                                        };
                                        const onUp = () => {
                                          window.removeEventListener(
                                            "mousemove",
                                            onMove
                                          );
                                          window.removeEventListener(
                                            "mouseup",
                                            onUp
                                          );
                                        };
                                        window.addEventListener(
                                          "mousemove",
                                          onMove
                                        );
                                        window.addEventListener(
                                          "mouseup",
                                          onUp
                                        );
                                      }}
                                    />
                                  )}
                                  {/* 右側拉伸 */}
                                  {isRightEdge && (
                                    <div
                                      style={{
                                        position: "absolute",
                                        right: 0,
                                        top: 0,
                                        width: 8,
                                        height: "100%",
                                        cursor: "ew-resize",
                                        background: "#333",
                                      }}
                                      title="Resize"
                                      onMouseDown={(e) => {
                                        e.preventDefault();
                                        const startX = e.clientX;
                                        const startDuration = task.duration;
                                        const onMove = (moveEvent) => {
                                          const diff = Math.round(
                                            (moveEvent.clientX - startX) / 40
                                          );
                                          onBarResize(
                                            task.id,
                                            startDuration + diff,
                                            "right"
                                          );
                                        };
                                        const onUp = () => {
                                          window.removeEventListener(
                                            "mousemove",
                                            onMove
                                          );
                                          window.removeEventListener(
                                            "mouseup",
                                            onUp
                                          );
                                        };
                                        window.addEventListener(
                                          "mousemove",
                                          onMove
                                        );
                                        window.addEventListener(
                                          "mouseup",
                                          onUp
                                        );
                                      }}
                                    />
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        )}
                      </Draggable>
                    ))}
                    {provided.placeholder}
                  </tbody>
                )}
              </Droppable>
            </DragDropContext>
          </table>
        </div>
      )}
      <div style={{ color: "#888", fontSize: 12, marginTop: 8 }}>
        Set project start/end and hours per day, add tasks (default or custom),
        drag bar to move, resize bar on both sides, select color above, select
        row and press Delete to remove, drag row to reorder, export/import
        Excel.
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62
      </div>
    </div>
  );
}
<<<<<<< HEAD

export default App; 
=======
>>>>>>> 9b77a01e5e89001776c44b40da27fe76796faf62

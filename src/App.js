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
  }
  return arr;
}

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
            }
          : t
      )
    );
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
  }

  // 鍵盤刪除
  React.useEffect(() => {
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
        <input
          type="number"
          min={1}
          max={allDateHours.length}
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
                  <th
                    key={i}
                    colSpan={hoursPerDay}
                    style={{
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
      </div>
    </div>
  );
}

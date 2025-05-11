import React, { useState } from "react";
import { Timeline, TimelineHeaders, DateHeader } from "react-calendar-timeline";
import moment from "moment";
import "react-calendar-timeline/lib/Timeline.css";
import { PlusIcon, TrashIcon, PencilIcon } from "@heroicons/react/24/outline";

const groups = [
  { id: 1, title: "任务组 1" },
  { id: 2, title: "任务组 2" },
  { id: 3, title: "任务组 3" },
];

const defaultItems = [
  {
    id: 1,
    group: 1,
    title: "任务 A",
    start_time: moment().add(-1, "day"),
    end_time: moment().add(1, "day"),
  },
  {
    id: 2,
    group: 2,
    title: "任务 B",
    start_time: moment().add(1, "day"),
    end_time: moment().add(3, "day"),
  },
];

export default function GanttChart() {
  const [items, setItems] = useState(defaultItems);
  const [newTaskTitle, setNewTaskTitle] = useState("");
  const [selectedGroup, setSelectedGroup] = useState(1);
  const [editingTask, setEditingTask] = useState(null);

  const handleItemMove = (itemId, dragTime, newGroupOrder) => {
    const updatedItems = items.map((item) =>
      item.id === itemId
        ? {
            ...item,
            start_time: moment(dragTime),
            end_time: moment(dragTime).add(item.end_time.diff(item.start_time)),
            group: groups[newGroupOrder].id,
          }
        : item
    );
    setItems(updatedItems);
  };

  const handleItemResize = (itemId, time, edge) => {
    const updatedItems = items.map((item) => {
      if (item.id === itemId) {
        return {
          ...item,
          start_time: edge === "left" ? moment(time) : item.start_time,
          end_time: edge === "right" ? moment(time) : item.end_time,
        };
      }
      return item;
    });
    setItems(updatedItems);
  };

  const addNewTask = () => {
    if (!newTaskTitle.trim()) return;

    const newTask = {
      id: items.length + 1,
      group: selectedGroup,
      title: newTaskTitle,
      start_time: moment(),
      end_time: moment().add(1, "day"),
    };

    setItems([...items, newTask]);
    setNewTaskTitle("");
  };

  const deleteTask = (taskId) => {
    setItems(items.filter((item) => item.id !== taskId));
  };

  const startEditing = (task) => {
    setEditingTask(task);
    setNewTaskTitle(task.title);
  };

  const saveEdit = () => {
    if (!editingTask || !newTaskTitle.trim()) return;

    const updatedItems = items.map((item) =>
      item.id === editingTask.id
        ? { ...item, title: newTaskTitle }
        : item
    );

    setItems(updatedItems);
    setEditingTask(null);
    setNewTaskTitle("");
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
      <div style={{ 
        display: 'flex', 
        gap: '1rem', 
        alignItems: 'center', 
        backgroundColor: 'white', 
        padding: '1rem',
        borderRadius: '0.5rem',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
      }}>
        <input
          type="text"
          value={newTaskTitle}
          onChange={(e) => setNewTaskTitle(e.target.value)}
          placeholder="输入任务名称"
          style={{
            flex: 1,
            padding: '0.5rem',
            border: '1px solid #e5e7eb',
            borderRadius: '0.25rem'
          }}
        />
        <select
          value={selectedGroup}
          onChange={(e) => setSelectedGroup(Number(e.target.value))}
          style={{
            padding: '0.5rem',
            border: '1px solid #e5e7eb',
            borderRadius: '0.25rem'
          }}
        >
          {groups.map((group) => (
            <option key={group.id} value={group.id}>
              {group.title}
            </option>
          ))}
        </select>
        {editingTask ? (
          <button
            onClick={saveEdit}
            style={{
              backgroundColor: '#10b981',
              color: 'white',
              padding: '0.5rem 1rem',
              borderRadius: '0.25rem',
              border: 'none',
              cursor: 'pointer'
            }}
          >
            保存编辑
          </button>
        ) : (
          <button
            onClick={addNewTask}
            style={{
              backgroundColor: '#3b82f6',
              color: 'white',
              padding: '0.5rem 1rem',
              borderRadius: '0.25rem',
              border: 'none',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '0.5rem'
            }}
          >
            <PlusIcon style={{ width: '1.25rem', height: '1.25rem' }} />
            添加任务
          </button>
        )}
      </div>

      <div style={{ 
        backgroundColor: 'white', 
        padding: '1rem',
        borderRadius: '0.5rem',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
      }}>
        <Timeline
          groups={groups}
          items={items}
          defaultTimeStart={moment().add(-3, "day")}
          defaultTimeEnd={moment().add(7, "day")}
          canMove
          canResize="both"
          onItemMove={handleItemMove}
          onItemResize={handleItemResize}
          itemRenderer={({ item }) => (
            <div style={{ position: 'relative' }}>
              <div style={{ padding: '0.5rem' }}>{item.title}</div>
              <div style={{ 
                position: 'absolute',
                right: '0.5rem',
                top: '0.5rem',
                display: 'none',
                gap: '0.5rem'
              }}>
                <button
                  onClick={() => startEditing(item)}
                  style={{
                    color: '#3b82f6',
                    background: 'none',
                    border: 'none',
                    cursor: 'pointer'
                  }}
                >
                  <PencilIcon style={{ width: '1rem', height: '1rem' }} />
                </button>
                <button
                  onClick={() => deleteTask(item.id)}
                  style={{
                    color: '#ef4444',
                    background: 'none',
                    border: 'none',
                    cursor: 'pointer'
                  }}
                >
                  <TrashIcon style={{ width: '1rem', height: '1rem' }} />
                </button>
              </div>
            </div>
          )}
        >
          <TimelineHeaders>
            <DateHeader unit="primaryHeader" />
            <DateHeader />
          </TimelineHeaders>
        </Timeline>
      </div>
    </div>
  );
} 
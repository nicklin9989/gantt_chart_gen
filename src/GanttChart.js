import React from "react";
import { Timeline, TimelineHeaders, DateHeader } from "react-calendar-timeline";
import moment from "moment";
import "react-calendar-timeline/lib/Timeline.css";

console.log({ Timeline, TimelineHeaders, DateHeader });

const groups = [
  { id: 1, title: "任务组 1" },
  { id: 2, title: "任务组 2" },
];

const items = [
  {
    id: 1,
    group: 1,
    title: "任务 A",
    start_time: moment().add(-1, "day"),
    end_time: moment().add(1, "day"),
  },
];

function GanttChart() {
  return (
    <div>
      <Timeline
        groups={groups}
        items={items}
        defaultTimeStart={moment().add(-3, "day")}
        defaultTimeEnd={moment().add(7, "day")}
      >
        <TimelineHeaders>
          <DateHeader unit="primaryHeader" />
          <DateHeader />
        </TimelineHeaders>
      </Timeline>
    </div>
  );
}

export default GanttChart;

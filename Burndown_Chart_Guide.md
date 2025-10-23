# Burndown Chart Feature Guide

## Overview
The Degrow Workflow Manager now includes a real-time burndown chart that visualizes workflow progress over time, showing the distribution of items by status across different dates.

## Features

### ✅ **Real-time Progress Visualization**
- **Live Updates**: Chart updates automatically when new entries are saved
- **Status Tracking**: Shows distribution of items by status (To Do, In Progress, Blocked, Done)
- **Time Series**: Displays progress over time with date-based grouping
- **Interactive**: Refresh button to manually update the chart

### ✅ **Visual Design**
- **Stacked Bar Chart**: Shows status distribution for each date
- **Color Coding**: Different colors for each status type
- **Professional Look**: Clean, modern chart design
- **Responsive**: Chart resizes with the window

### ✅ **Data Insights**
- **Total Items**: Shows total number of items in the system
- **Status Breakdown**: Visual representation of workflow status distribution
- **Trend Analysis**: See how workflow progresses over time
- **Performance Metrics**: Track completion rates and bottlenecks

## Chart Layout

### **Window Design**
```
┌─────────────────────────────────────────────────────────────────────────┐
│ Degrow Workflow Manager                    Progress Burndown Chart      │
├─────────────────────────────────────────────────────────────────────────┤
│ Form Fields (Left Side)                 │ Chart Section (Right Side)   │
│ ┌─────────────────────────────────────┐ │ ┌─────────────────────────┐ │
│ │ State: [Dropdown]                   │ │ │ [Refresh Chart]         │ │
│ │ Stage: [Dropdown]                   │ │ │                         │ │
│ │ Status: [Dropdown]                  │ │ │ ┌─────────────────────┐ │ │
│ │ CLLI: [Text] [Dropdown]             │ │ │ │                     │ │ │
│ │ City: [Text Field]                  │ │ │ │   Burndown Chart    │ │ │
│ │ LATA: [Text Field]                  │ │ │ │   Visualization     │ │ │
│ │ Equipment: [Dropdown]               │ │ │ │                     │ │ │
│ │ Milestone: [Auto-calculated]       │ │ │ │                     │ │ │
│ │                                     │ │ │ │                     │ │ │
│ │ [Save Entry] [Clear Form] [Exit]   │ │ │ │                     │ │ │
│ │                                     │ │ │ │                     │ │ │
│ │ Status: Ready                       │ │ │ │                     │ │ │
│ │                                     │ │ │ │                     │ │ │
│ │ Recent Entries:                     │ │ │ │                     │ │ │
│ │ ┌─────────────────────────────────┐ │ │ │ └─────────────────────┘ │ │
│ │ │ Entry 1                         │ │ │                         │ │
│ │ │ Entry 2                         │ │ │                         │ │
│ │ │ ...                             │ │ │                         │ │
│ │ └─────────────────────────────────┘ │ │                         │ │
│ └─────────────────────────────────────┘ │                         │ │
└─────────────────────────────────────────────────────────────────────────┘
```

## Chart Features

### **Visual Elements**
- **Stacked Bars**: Each date shows a stacked bar with different statuses
- **Color Legend**: 
  - 🔴 **To Do**: Red (#ff6b6b)
  - 🔵 **In Progress**: Teal (#4ecdc4)
  - 🟠 **Blocked**: Orange (#ffa726)
  - 🟢 **Done**: Green (#66bb6a)
- **Grid Lines**: Subtle grid for easier reading
- **Total Count**: Shows total items in the system

### **Data Processing**
- **Date Grouping**: Groups entries by creation date
- **Status Counting**: Counts items by status for each date
- **Time Series**: Orders data chronologically
- **Real-time Updates**: Refreshes when new entries are added

## How to Use

### **Automatic Updates**
- **Save Entry**: Chart updates automatically when you save a new entry
- **Real-time**: No manual refresh needed for new data
- **Live Data**: Always shows current database state

### **Manual Refresh**
- **Refresh Button**: Click "Refresh Chart" to manually update
- **Force Update**: Useful if data changes outside the application
- **Error Recovery**: Refresh if chart encounters errors

### **Chart Interaction**
- **Zoom**: Chart automatically adjusts to data range
- **Legend**: Click legend items to show/hide status types
- **Tooltips**: Hover over bars for detailed information

## Data Visualization

### **What the Chart Shows**
1. **Daily Progress**: How many items were created each day
2. **Status Distribution**: Breakdown of items by status
3. **Workflow Trends**: Patterns in workflow progression
4. **Bottlenecks**: Areas where items get stuck (Blocked status)
5. **Completion Rate**: How many items reach "Done" status

### **Chart Interpretation**
- **Growing Bars**: More items being created over time
- **Color Distribution**: Balance between different statuses
- **Trend Lines**: Overall workflow direction
- **Peak Days**: Days with highest activity
- **Status Shifts**: Movement of items through workflow stages

## Technical Details

### **Data Source**
- **Database**: Reads from BASHFlowSandbox table
- **Filtering**: Only includes "Degrow Workflow Entry" records
- **Grouping**: Groups by DateCreated and Status
- **Sorting**: Orders by date for chronological display

### **Chart Technology**
- **Matplotlib**: Professional charting library
- **Tkinter Integration**: Embedded in GUI
- **Real-time Updates**: Canvas refresh on data changes
- **Error Handling**: Graceful handling of data issues

### **Performance**
- **Efficient Queries**: Optimized database queries
- **Memory Management**: Chart objects properly managed
- **Smooth Updates**: Fast chart refresh
- **Responsive**: Chart updates don't block UI

## Installation Requirements

### **Python Packages**
```bash
pip install matplotlib numpy pandas openpyxl
```

Or install from requirements file:
```bash
pip install -r requirements.txt
```

### **System Requirements**
- **Python 3.8+**: Required for matplotlib
- **Display**: GUI display required for charts
- **Memory**: Additional memory for chart rendering
- **Graphics**: Hardware acceleration recommended

## Troubleshooting

### **Common Issues**

#### 1. "matplotlib not found" error
**Solution**: Install matplotlib
```bash
pip install matplotlib
```

#### 2. Chart not displaying
**Possible Causes**:
- Display issues (headless environment)
- Matplotlib backend problems
- Memory issues

**Solution**:
- Ensure GUI display is available
- Check matplotlib installation
- Restart application

#### 3. Chart not updating
**Possible Causes**:
- Database connection issues
- Chart refresh problems
- Data format issues

**Solution**:
- Check database connectivity
- Use refresh button
- Check console for errors

#### 4. Performance issues
**Possible Causes**:
- Large datasets
- Memory constraints
- Slow database queries

**Solution**:
- Limit data range if needed
- Check database performance
- Monitor memory usage

### **Debug Information**
Check console output for:
- Chart creation status
- Data loading results
- Error messages
- Performance metrics

## Customization

### **Chart Appearance**
```python
# Modify colors
colors = {'To Do': '#ff6b6b', 'In Progress': '#4ecdc4', 
          'Blocked': '#ffa726', 'Done': '#66bb6a'}

# Change chart size
self.fig = Figure(figsize=(6, 4), dpi=100)

# Modify title
self.ax.set_title('Workflow Progress Over Time', fontsize=14, fontweight='bold')
```

### **Data Filtering**
```python
# Filter by date range
cursor.execute("""
    SELECT DateCreated, Status, COUNT(*) as Count
    FROM BASHFlowSandbox 
    WHERE Description = 'Degrow Workflow Entry'
    AND DateCreated >= '2024-01-01'
    GROUP BY DateCreated, Status
    ORDER BY DateCreated
""")
```

## Integration Notes

### **VBA Compatibility**
- **No Impact**: Chart doesn't affect VBA integration
- **Database Shared**: Uses same database as VBA system
- **Data Consistency**: Chart reflects all database changes

### **Performance Impact**
- **Minimal**: Chart updates are lightweight
- **Non-blocking**: Chart updates don't block UI
- **Efficient**: Optimized database queries

## Future Enhancements

### **Potential Improvements**
- **Interactive Charts**: Click to drill down into data
- **Export Functionality**: Save charts as images
- **Multiple Chart Types**: Line charts, pie charts
- **Advanced Analytics**: Trend analysis, forecasting
- **Custom Time Ranges**: Filter by date ranges
- **Status Transitions**: Track item movement between statuses

### **Advanced Features**
- **Real-time Updates**: WebSocket-based live updates
- **Historical Analysis**: Compare different time periods
- **Performance Metrics**: Velocity, throughput analysis
- **Predictive Analytics**: Forecast completion dates

The burndown chart provides valuable insights into workflow progress, helping teams understand their productivity patterns and identify areas for improvement!

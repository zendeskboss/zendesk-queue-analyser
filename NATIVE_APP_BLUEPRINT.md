**Mission:**
This app is a top navigation bar (nav_bar) analytics dashboard that provides deep insights into queue performance, including wait times, queue sizes, and raw event data across Support, Talk, and Messaging channels, helping admins and team leads monitor and optimize routing and staffing.

**Action:**
Create a React-based app using the Zendesk App Framework (ZAF) that fetches data natively from the `/api/v2/queues` and `/api/v2/queue_events` endpoints. It must allow the user to select a date range (enforcing a maximum limit of 90 days), specific channels, and multiple queues. Upon clicking "Run analysis", it should fetch all paginated queue events for the selected parameters (must handle cursor pagination using `links.next` with `page_size=100`), process the metrics (calculating wait times based on `start_queue_time` and `end_queue_time` for outbound events, snapshot sizes, and hourly averages), and display the aggregated data.

**Parts:**
* Use Zendesk Garden components extensively for the UI (Tabs, Tables, Combobox Dropdowns, DatePickerRange, Alerts, Buttons, Spinners) to ensure it looks native to the Zendesk interface.
* Build an "Analysis Parameters" collapsible panel at the top containing the DatePickerRange, channel selection chips (Messaging, Talk, Support), a multi-select Combobox for filtering specific queues, and a "Run analysis" button.
* Create three main Tabs: "Wait Times", "Queue Size", and "Raw Events".
* In the "Wait Times" tab: Include a Summary Card displaying high-level metrics (Avg Wait, Max Wait, Min Wait, Interactions, Unique Tickets). Add a `recharts` ComposedChart showing "Daily Trend — Volume vs Wait Time" (using Bars for volume and Lines for wait times). Include ZAF DataTables breaking down wait times By Queue, By Channel, and By Day.
* In the "Queue Size" tab: Include a Summary Card for snapshot metrics. Crucially, add a Garden Warning Alert that triggers if all snapshot values return as zero—this must explain that it is a known Zendesk limitation for the Support channel (where `current_channel_queue_size` is not populated). Include a `recharts` ComposedChart for "Daily Trend — Queue Size" and DataTables for sizes By Queue, By Channel, and By Day.
* In the "Raw Events" tab: Display a paginated, sortable, and searchable Garden Table of the raw queue events. Include a functional "Export to CSV" button that downloads the table data.
* Implement robust error handling with Garden Alerts for API failures or date validation errors, and use Spinners to indicate loading states during initial queue fetching and heavy queue event data processing.

**Scope:**
The app is strictly for the `nav_bar` location. It is intended for administrators and supervisors analyzing internal queue performance data within the Zendesk account.
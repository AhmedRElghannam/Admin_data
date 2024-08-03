# 📊 Sales Data Management for Company Administration

## 🛠️ Requirements
1. **Control Unit**: For the data administrator to process the data.
2. **Manager's Dashboard**: To view overall performance metrics.
3. **Team Leaders' Dashboards**: To review their respective metrics.

**Data Files Type**: CSV

---

## 📈 Step 1: Data Analysis
**Objective**: Identify the key questions that the dashboards need to answer.  
**Action**: Data was analyzed using Jupyter and Python with various libraries. [View the code here](#).

---

## 📊 Step 2: Selecting the Appropriate Tool for Sharing Analyzed Data with the Manager
**Objective**: Create a dashboard for the manager.  
**Action**: Google Looker Studio was selected to create the dashboard by uploading the data to Google Sheets and linking it to the dashboard.

---

## 📋 Step 3: Data Processing for Sorting by Team Leaders and Displaying to Them
**Objective**: Display data relevant to each team.  
**Action**: Company data was aggregated based on team leaders as indicated in the data files. Additional data was appended, and a separate dashboard was created for each team leader to display their sales data, all linked to the same Google Sheet.

---

## 🖥️ Step 4: Structuring Work for the Data Administrator and Implementing the Processing in Jupyter as Code for Daily Operations
**Objective**: Simplify daily operations for the data administrator.  
**Action**: Streamlit was used to build a control panel for the data administrator to upload daily data files and connect Streamlit to Google Cloud via the API to update the Google Sheet. Additionally, data processing tasks, such as handling complaints and sending emails to team leaders and the manager, were automated.

---

## ✅ Step 5: Testing Streamlit Code on the Local Server and Creating the Requirements File
**Objective**: Test the code and ensure it runs on the server.  
**Action**: The code was successfully tested, and [you can view the source code on GitHub](#). A requirements file was created and the app was deployed on the free Streamlit server for practical use.

---

## 🔗 Useful Links:
- **Manager's Dashboard**: [Google Looker Studio](#)
- **Team Leaders' Dashboard**: [Google Looker Studio](#)
- **GitHub Repository**: [GitHub](#)
- **Data Administrator's Control Panel**: [Streamlit App](#)

---

Feel free to ask if you need any more improvements or specific details!

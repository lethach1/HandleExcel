import React, { useState } from 'react';
import { Tabs } from 'antd';
import './App.css';
import MainScreen from './components/MainScreen';
import DataTable from './components/DataTable';

const { TabPane } = Tabs;

function App() {
  return (
    <div className="app-container">
      <Tabs defaultActiveKey="1">
        <TabPane tab="Excel Import/Export" key="1">
          <MainScreen />
        </TabPane>
        <TabPane tab="Data Viewer" key="2">
          <DataTable />
        </TabPane>
      </Tabs>
    </div>
  );
}

export default App;

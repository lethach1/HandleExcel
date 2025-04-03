import React from 'react';
import { Tabs, ConfigProvider, theme } from 'antd';
import './App.css';
import MainScreen from './components/MainScreen';
import DataTable from './components/DataTable';
import { FileExcelOutlined, TableOutlined } from '@ant-design/icons';

const { TabPane } = Tabs;

function App() {
  return (
    <ConfigProvider
      theme={{
        algorithm: theme.defaultAlgorithm,
        token: {
          colorPrimary: '#1890ff',
          borderRadius: 8,
          colorBgContainer: '#ffffff',
        },
      }}
    >
      <div className="app-container">
        <Tabs 
          defaultActiveKey="1" 
          size="large"
          centered
          tabBarStyle={{ fontWeight: 'bold' }}
        >
          <TabPane 
            tab={
              <span>
                <FileExcelOutlined style={{ marginRight: 8 }} />
                Excel Import/Export
              </span>
            } 
            key="1"
          >
            <MainScreen />
          </TabPane>
          <TabPane 
            tab={
              <span>
                <TableOutlined style={{ marginRight: 8 }} />
                Data Viewer
              </span>
            } 
            key="2"
          >
            <DataTable />
          </TabPane>
        </Tabs>
      </div>
    </ConfigProvider>
  );
}

export default App;

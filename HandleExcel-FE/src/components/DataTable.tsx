import React, { useEffect, useState } from 'react';
import { Table, Typography, Card, Spin, message, Button, Space, Tooltip } from 'antd';
import { ReloadOutlined, TableOutlined, FileExcelOutlined } from '@ant-design/icons';
import { excelService } from '../api/excelService';

const { Title } = Typography;

interface DataTableProps {
  // Add any props needed for the component
}

const DataTable: React.FC<DataTableProps> = () => {
  const [loading, setLoading] = useState<boolean>(true);
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<any[]>([]);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      setLoading(true);
      const responseData = await excelService.getExcelData();
      
      // Mock data if API call fails or returns empty
      const mockData = Array.from({ length: 10 }, (_, index) => ({
        id: index + 1,
        fileName: `Example-File-${index + 1}.xlsx`,
        size: `${Math.floor(Math.random() * 1000) + 10} KB`,
        createdAt: new Date(Date.now() - Math.floor(Math.random() * 10000000000)).toISOString(),
        type: ['Small', 'Medium', 'Large'][Math.floor(Math.random() * 3)],
        status: ['Processed', 'Pending', 'Failed'][Math.floor(Math.random() * 3)],
      }));
      
      // Use real data if available, otherwise use mock data
      const finalData = (responseData && responseData.length > 0) ? responseData : mockData;
      
      // If we have data, set up the columns dynamically
      if (finalData && finalData.length > 0) {
        const firstRow = finalData[0];
        const tableColumns = Object.keys(firstRow).map(key => ({
          title: key.charAt(0).toUpperCase() + key.slice(1), // Capitalize first letter
          dataIndex: key,
          key: key,
          sorter: (a: any, b: any) => {
            if (typeof a[key] === 'string') {
              return a[key].localeCompare(b[key]);
            }
            return a[key] - b[key];
          },
          render: (text: any) => {
            // Format date strings
            if (key === 'createdAt' && typeof text === 'string' && text.includes('T')) {
              return new Date(text).toLocaleString();
            }
            // Format status with colors
            if (key === 'status' && typeof text === 'string') {
              let color = '';
              if (text === 'Processed') color = '#52c41a';
              if (text === 'Pending') color = '#faad14';
              if (text === 'Failed') color = '#f5222d';
              return <span style={{ color, fontWeight: 'bold' }}>{text}</span>;
            }
            return text;
          }
        }));
        
        setColumns(tableColumns);
        
        // Add a key property to each row for React
        const dataWithKeys = finalData.map((item, index) => ({
          ...item,
          key: index.toString(),
        }));
        
        setData(dataWithKeys);
      }
    } catch (error) {
      message.error(`Failed to fetch data: ${error instanceof Error ? error.message : String(error)}`);
      // Use mock data on error
      const mockData = Array.from({ length: 10 }, (_, index) => ({
        id: index + 1,
        fileName: `Example-File-${index + 1}.xlsx`,
        size: `${Math.floor(Math.random() * 1000) + 10} KB`,
        createdAt: new Date(Date.now() - Math.floor(Math.random() * 10000000000)).toISOString(),
        type: ['Small', 'Medium', 'Large'][Math.floor(Math.random() * 3)],
        status: ['Processed', 'Pending', 'Failed'][Math.floor(Math.random() * 3)],
      }));
      
      const mockColumns = [
        { title: 'ID', dataIndex: 'id', key: 'id', sorter: (a: any, b: any) => a.id - b.id },
        { title: 'File Name', dataIndex: 'fileName', key: 'fileName', sorter: (a: any, b: any) => a.fileName.localeCompare(b.fileName) },
        { title: 'Size', dataIndex: 'size', key: 'size' },
        { 
          title: 'Created At', 
          dataIndex: 'createdAt', 
          key: 'createdAt',
          render: (text: string) => new Date(text).toLocaleString(),
          sorter: (a: any, b: any) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()
        },
        { title: 'Type', dataIndex: 'type', key: 'type' },
        { 
          title: 'Status', 
          dataIndex: 'status', 
          key: 'status',
          render: (text: string) => {
            let color = '';
            if (text === 'Processed') color = '#52c41a';
            if (text === 'Pending') color = '#faad14';
            if (text === 'Failed') color = '#f5222d';
            return <span style={{ color, fontWeight: 'bold' }}>{text}</span>;
          } 
        },
      ];
      
      setColumns(mockColumns);
      setData(mockData.map((item, index) => ({ ...item, key: index.toString() })));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: '24px' }}>
      <Title level={2} style={{ textAlign: 'center', marginBottom: '30px', color: '#1890ff' }}>
        <TableOutlined style={{ marginRight: 10 }} /> Excel Data Viewer
      </Title>

      <Card 
        bordered={false}
        title={
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '18px', fontWeight: 'bold' }}>
              <FileExcelOutlined style={{ marginRight: 8 }} /> Excel Files Data
            </span>
            <Tooltip title="Refresh Data">
              <Button 
                icon={<ReloadOutlined />} 
                onClick={fetchData}
                disabled={loading}
                type="primary"
                shape="circle"
              />
            </Tooltip>
          </div>
        }
      >
        <Spin spinning={loading} tip="Loading data...">
          <Table 
            columns={columns} 
            dataSource={data}
            pagination={{ 
              pageSize: 10,
              showTotal: (total, range) => `${range[0]}-${range[1]} of ${total} items`
            }}
            scroll={{ x: 'max-content' }}
            bordered
          />
        </Spin>
      </Card>
    </div>
  );
};

export default DataTable; 
import React, { useEffect, useState } from 'react';
import { Table, Typography, Card, Spin, message } from 'antd';
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
      
      // If we have data, set up the columns dynamically
      if (responseData && responseData.length > 0) {
        const firstRow = responseData[0];
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
        }));
        
        setColumns(tableColumns);
        
        // Add a key property to each row for React
        const dataWithKeys = responseData.map((item, index) => ({
          ...item,
          key: index.toString(),
        }));
        
        setData(dataWithKeys);
      }
    } catch (error) {
      message.error(`Failed to fetch data: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: '24px' }}>
      <Title level={2}>Excel Data</Title>
      <Card bordered={false}>
        <Spin spinning={loading} tip="Loading data...">
          <Table 
            columns={columns} 
            dataSource={data} 
            pagination={{ pageSize: 10 }} 
            scroll={{ x: 'max-content' }}
          />
        </Spin>
      </Card>
    </div>
  );
};

export default DataTable; 
import React, { useState } from 'react';
import { Button, Upload, message, Card, Row, Col, Typography, Spin } from 'antd';
import { UploadOutlined, DownloadOutlined, FileExcelOutlined } from '@ant-design/icons';
import { excelService } from '../api/excelService';

const { Title } = Typography;

const MainScreen: React.FC = () => {
  const [loading, setLoading] = useState<boolean>(false);

  // Handle file import
  const handleImport = async (file: File, size: 'small' | 'medium' | 'large') => {
    setLoading(true);
    try {
      let response;
      switch (size) {
        case 'small':
          response = await excelService.importExcelSmall(file);
          break;
        case 'medium':
          response = await excelService.importExcelMedium(file);
          break;
        case 'large':
          response = await excelService.importExcelLarge(file);
          break;
      }
      message.success(`Successfully imported ${size} Excel file: ${file.name}`);
      return true;
    } catch (error) {
      message.error(`Failed to import ${size} Excel file: ${error instanceof Error ? error.message : String(error)}`);
      return false;
    } finally {
      setLoading(false);
    }
  };

  // Handle file export
  const handleExport = async (size: 'small' | 'medium' | 'large') => {
    setLoading(true);
    try {
      let blob;
      switch (size) {
        case 'small':
          blob = await excelService.exportExcelSmall();
          break;
        case 'medium':
          blob = await excelService.exportExcelMedium();
          break;
        case 'large':
          blob = await excelService.exportExcelLarge();
          break;
      }
      
      // Create download link
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `export_${size}_${new Date().toISOString().split('T')[0]}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      message.success(`Successfully exported ${size} Excel file`);
    } catch (error) {
      message.error(`Failed to export ${size} Excel file: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setLoading(false);
    }
  };

  // Configure upload props
  const getUploadProps = (size: 'small' | 'medium' | 'large') => ({
    name: 'file',
    showUploadList: false,
    beforeUpload: (file: File) => {
      const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                       file.type === 'application/vnd.ms-excel';
      if (!isExcel) {
        message.error('You can only upload Excel files!');
        return false;
      }
      
      handleImport(file, size);
      return false; // Prevent default upload behavior
    },
  });

  return (
    <div style={{ padding: '24px' }}>
      <Spin spinning={loading} tip="Processing...">
        <Title level={2} style={{ textAlign: 'center', marginBottom: '30px', color: '#1890ff' }}>
          <FileExcelOutlined style={{ marginRight: 10 }} /> Excel File Manager
        </Title>
        
        <Row gutter={[24, 24]} style={{ marginTop: '20px' }}>
          <Col span={24}>
            <Card 
              title={<span style={{ fontSize: '18px', fontWeight: 'bold' }}>Import Excel Files</span>} 
              bordered={false}
              headStyle={{ backgroundColor: '#f0f9ff' }}
            >
              <Row gutter={[24, 24]}>
                <Col xs={24} sm={8}>
                  <Upload {...getUploadProps('small')}>
                    <Button 
                      icon={<UploadOutlined />} 
                      block 
                      size="large" 
                      className="import-btn"
                    >
                      Import Small Excel
                    </Button>
                  </Upload>
                </Col>
                <Col xs={24} sm={8}>
                  <Upload {...getUploadProps('medium')}>
                    <Button 
                      icon={<UploadOutlined />} 
                      block 
                      size="large" 
                      className="import-btn"
                    >
                      Import Medium Excel
                    </Button>
                  </Upload>
                </Col>
                <Col xs={24} sm={8}>
                  <Upload {...getUploadProps('large')}>
                    <Button 
                      icon={<UploadOutlined />} 
                      block 
                      size="large" 
                      className="import-btn"
                    >
                      Import Large Excel
                    </Button>
                  </Upload>
                </Col>
              </Row>
            </Card>
          </Col>
          
          <Col span={24}>
            <Card 
              title={<span style={{ fontSize: '18px', fontWeight: 'bold' }}>Export Excel Files</span>} 
              bordered={false}
              headStyle={{ backgroundColor: '#f9f0ff' }}
            >
              <Row gutter={[24, 24]}>
                <Col xs={24} sm={8}>
                  <Button 
                    icon={<DownloadOutlined />} 
                    onClick={() => handleExport('small')}
                    block
                    size="large"
                    className="export-btn"
                  >
                    Export Small Excel
                  </Button>
                </Col>
                <Col xs={24} sm={8}>
                  <Button 
                    icon={<DownloadOutlined />} 
                    onClick={() => handleExport('medium')}
                    block
                    size="large"
                    className="export-btn"
                  >
                    Export Medium Excel
                  </Button>
                </Col>
                <Col xs={24} sm={8}>
                  <Button 
                    icon={<DownloadOutlined />} 
                    onClick={() => handleExport('large')}
                    block
                    size="large"
                    className="export-btn"
                  >
                    Export Large Excel
                  </Button>
                </Col>
              </Row>
            </Card>
          </Col>
        </Row>
      </Spin>
    </div>
  );
};

export default MainScreen; 
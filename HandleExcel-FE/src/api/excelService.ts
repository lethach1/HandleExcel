import axios from 'axios';

// Base API URL - change to your actual backend URL
const API_URL = 'http://localhost:3000/api';

// Types for request and response
export interface ExcelFile {
  id: string;
  name: string;
  size: string;
  createdAt: string;
  // Add any other fields relevant to your Excel files
}

// API functions for the Excel operations
export const excelService = {
  // Import Excel file functions
  importExcelSmall: async (file: File): Promise<ExcelFile> => {
    const formData = new FormData();
    formData.append('file', file);
    const response = await axios.post(`${API_URL}/excel/import/small`, formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    });
    return response.data;
  },

  importExcelMedium: async (file: File): Promise<ExcelFile> => {
    const formData = new FormData();
    formData.append('file', file);
    const response = await axios.post(`${API_URL}/excel/import/medium`, formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    });
    return response.data;
  },

  importExcelLarge: async (file: File): Promise<ExcelFile> => {
    const formData = new FormData();
    formData.append('file', file);
    const response = await axios.post(`${API_URL}/excel/import/large`, formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    });
    return response.data;
  },

  // Export Excel file functions
  exportExcelSmall: async (): Promise<Blob> => {
    const response = await axios.get(`${API_URL}/excel/export/small`, {
      responseType: 'blob'
    });
    return response.data;
  },

  exportExcelMedium: async (): Promise<Blob> => {
    const response = await axios.get(`${API_URL}/excel/export/medium`, {
      responseType: 'blob'
    });
    return response.data;
  },

  exportExcelLarge: async (): Promise<Blob> => {
    const response = await axios.get(`${API_URL}/excel/export/large`, {
      responseType: 'blob'
    });
    return response.data;
  },

  // Get Excel file data for the table display
  getExcelData: async (): Promise<any[]> => {
    const response = await axios.get(`${API_URL}/excel/data`);
    return response.data;
  }
}; 
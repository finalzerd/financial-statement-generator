import React from 'react';
import { generateSampleCSV } from '../utils/csvSampleGenerator';

const SampleFileDownloader: React.FC = () => {
  
  const downloadCSVSample = async () => {
    try {
      const csvContent = generateSampleCSV();
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.href = url;
      link.download = 'TB01_Sample.csv';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Error downloading CSV sample:', error);
      alert('เกิดข้อผิดพลาดในการดาวน์โหลดไฟล์ตัวอย่าง');
    }
  };

  return (
    <div className="sample-downloader">
      <h3>📥 ดาวน์โหลดไฟล์ตัวอย่าง</h3>
      
      <div className="download-buttons">
        <button 
          className="download-btn csv"
          onClick={downloadCSVSample}
        >
          📄 ไฟล์ CSV ตัวอย่าง
        </button>
      </div>
      
      <div className="sample-info">
        <div className="sample-description">
          <h4>ไฟล์ CSV ตัวอย่าง:</h4>
          <ul>
            <li>รูปแบบ CSV ที่ระบุข้อมูลงบทดลอง</li>
            <li>ระบบจะให้กรอกข้อมูลบริษัทเพิ่มเติม</li>
            <li>สามารถสร้างงบปีเดียวหรือหลายปีได้</li>
            <li>ยอดยกมาต้นงวด = ปีก่อน, ยอดยกมางวดนี้ = ปีปัจจุบัน</li>
            <li>หากมีข้อมูลในคอลัมน์ยอดยกมาต้นงวด = งบหลายปี</li>
          </ul>
        </div>
      </div>
      
      <div className="file-structure-info">
        <h4>🏗️ โครงสร้างไฟล์ที่ต้องการ:</h4>
        <div className="structure-grid">
          <div className="structure-item">
            <h5>คอลัมน์ที่จำเป็น</h5>
            <ul>
              <li>ชื่อบัญชี</li>
              <li>รหัสบัญชี</li>
              <li>ยอดยกมาต้นงวด (ปีก่อน)</li>
              <li>ยอดยกมางวดนี้ (ปีปัจจุบัน)</li>
              <li>เดบิต</li>
              <li>เครดิต</li>
            </ul>
          </div>
          
          <div className="structure-item">
            <h5>ข้อมูลบริษัท</h5>
            <ul>
              <li>ชื่อบริษัท</li>
              <li>ประเภทบริษัท</li>
              <li>ที่อยู่</li>
              <li>ปีบัญชี</li>
              <li>จะกรอกผ่านฟอร์ม</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SampleFileDownloader;

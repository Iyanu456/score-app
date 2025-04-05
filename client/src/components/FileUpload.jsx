import { useState } from 'react';
import axios from 'axios';

export default function FileUpload() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [successMessage, setSuccessMessage] = useState('');

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
    setError('');
    setSuccessMessage(''); // Clear success message when a new file is selected
  };

  const handleUpload = async () => {
    if (!file) return;

    const formData = new FormData();
    formData.append('document', file);

    try {
      setLoading(true);
      setError('');
      const res = await axios.post('/api/upload', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob', // Important to handle binary data (file)
      });

      // Trigger the download
      handleDownload(res.data);
    } catch (err) {
      console.error(err);
      setError(err.response?.data?.error || 'Something went wrong');
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = (fileBlob) => {
    const link = document.createElement('a');
    const url = window.URL.createObjectURL(fileBlob);
    link.href = url;
    link.setAttribute('download', 'processed-file.docx');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Set success message
    setSuccessMessage('File downloaded successfully!');
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-center flex-col sm:flex-row items-center gap-4">
        <input
          type="file"
          accept=".docx"
          onChange={handleFileChange}
          className="border-2 border-dashed border-indigo-300 px-4 py-2 rounded-md w-full sm:w-auto"
        />
        <button
          onClick={handleUpload}
          disabled={loading}
          className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2 rounded-md transition disabled:opacity-50"
        >
          {loading ? 'Processing...' : 'Upload & Extract'}
        </button>
      </div>

      {error && <p className="text-red-600 font-semibold">{error}</p>}
      {successMessage && <p className="text-green-600 font-semibold">{successMessage}</p>} {/* Display success message */}
    </div>
  );
}

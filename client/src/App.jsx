import FileUpload from "./components/FileUpload";

function App() {
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-100 to-indigo-100 flex flex-col items-center justify-center px-4">
      <div className="bg-white shadow-xl rounded-2xl p-8 w-full max-w-3xl">
        <h1 className="text-3xl font-bold text-indigo-700 mb-2 text-center">
          📄 Student Results Extractor
        </h1>
        <p className="text-gray-600 text-center mb-6">
          Upload a DOCX file to extract structured student data
        </p>
        <FileUpload />
      </div>
    </div>
  );
}

export default App;

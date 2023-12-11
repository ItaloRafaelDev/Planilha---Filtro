import React from 'react';
import './App.css'; // Se houver um arquivo CSS global

import UploadExcel from './components/Excel/UploadExcel';

function App() {
  return (
    <div className="App">
      <main>
        <UploadExcel />
      </main>
      <footer>
        {/* Adicione aqui o conteúdo do rodapé, se necessário */}
      </footer>
    </div>
  );
}

export default App;

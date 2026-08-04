import React, { useState } from 'react';
import { DrawerContato } from './components/DrawerContato';

export default function App() {
  // Estado para armazenar a lista de lojas (exemplo de dados iniciais)
  const [lojas, setLojas] = useState([
    {
      pvAbadi: '1001',
      razaoSocial: 'Posto AmPm Central',
      statusContato: 'A Contatar',
      telefoneContato: '11999998888',
      nomeContato: 'Carlos Gerente',
      latitude: -23.55052,
      longitude: -46.633308
    },
    {
      pvAbadi: '1002',
      razaoSocial: 'Posto AmPm Marginal',
      statusContato: 'A Contatar',
      telefoneContato: '11977776666',
      nomeContato: 'Ana Supervisora',
      latitude: -23.56168,
      longitude: -46.655981
    }
  ]);

  // Estados de controle do Drawer
  const [lojaSelecionada, setLojaSelecionada] = useState(null);
  const [isDrawerOpen, setIsDrawerOpen] = useState(false);

  // Funções de Ação
  const handleAbrirDrawer = (loja) => {
    setLojaSelecionada(loja);
    setIsDrawerOpen(true);
  };

  const handleSalvarContato = (pvAbadi, dadosAtualizados) => {
    setLojas((prev) =>
      prev.map((item) =>
        item.pvAbadi === pvAbadi ? { ...item, ...dadosAtualizados } : item
      )
    );
    setIsDrawerOpen(false);
  };

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h1 style={{ color: '#002060' }}>Painel Tático AmPm</h1>

      {/* Tabela de Lojas */}
      <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '20px' }}>
        <thead>
          <tr style={{ backgroundColor: '#002060', color: '#fff', textAlign: 'left' }}>
            <th style={{ padding: '10px' }}>PV Abadi</th>
            <th style={{ padding: '10px' }}>Razão Social</th>
            <th style={{ padding: '10px' }}>Status Contato</th>
            <th style={{ padding: '10px' }}>Ações</th>
          </tr>
        </thead>
        <tbody>
          {lojas.map((loja) => (
            <tr key={loja.pvAbadi} style={{ borderBottom: '1px solid #ddd' }}>
              <td style={{ padding: '10px' }}>{loja.pvAbadi}</td>
              <td style={{ padding: '10px' }}>{loja.razaoSocial}</td>
              <td style={{ padding: '10px' }}>{loja.statusContato}</td>
              <td style={{ padding: '10px' }}>
                <button
                  onClick={() => handleAbrirDrawer(loja)}
                  style={{
                    backgroundColor: '#e0a96d',
                    border: 'none',
                    padding: '6px 12px',
                    borderRadius: '4px',
                    cursor: 'pointer',
                    fontWeight: 'bold'
                  }}
                >
                  Registrar Contato
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* Componente Drawer Lateral */}
      <DrawerContato
        loja={lojaSelecionada}
        isOpen={isDrawerOpen}
        onClose={() => setIsDrawerOpen(false)}
        onSave={handleSalvarContato}
      />
    </div>
  );
}

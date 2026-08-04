import React, { useState, useEffect } from 'react';

export const DrawerContato = ({ loja, isOpen, onClose, onSave, instrutoresMaisProximos }) => {
  const [formData, setFormData] = useState({
    nomeContato: '',
    cargoContato: '',
    telefoneContato: '',
    emailContato: '',
    canal: 'Ligação',
    statusContato: 'A Contatar',
    dataAgendamento: '',
    motivoRecusa: '',
    observacoes: '',
  });

  const [rascunhoSalvo, setRascunhoSalvo] = useState(false);

  useEffect(() => {
    if (loja) {
      setFormData({
        nomeContato: loja.nomeContato || '',
        cargoContato: loja.cargoContato || 'Gerente',
        telefoneContato: loja.telefoneContato || '',
        emailContato: loja.emailContato || '',
        canal: loja.canal || 'Ligação',
        statusContato: loja.statusContato || 'A Contatar',
        dataAgendamento: loja.dataAgendamento || '',
        motivoRecusa: loja.motivoRecusa || '',
        observacoes: loja.observacoes || '',
      });
    }
  }, [loja]);

  if (!isOpen || !loja) return null;

  const handleChange = (field, value) => {
    setFormData((prev) => ({ ...prev, [field]: value }));
    setRascunhoSalvo(true);
    setTimeout(() => setRascunhoSalvo(false), 2000);
  };

  const handleOpenWhatsApp = () => {
    const numLimpo = formData.telefoneContato.replace(/\D/g, '');
    if (!numLimpo) return alert('Por favor, insira um número de telefone válido.');
    const texto = encodeURIComponent(`Olá! Sou da equipe de treinamento AmPm. Gostaria de falar sobre o agendamento do retreinamento da unidade ${loja.razaoSocial || loja.pvAbadi}.`);
    window.open(`https://wa.me/55${numLimpo}?text=${texto}`, '_blank');
  };

  const badgeColors = {
    'A Contatar': '#6c757d',
    'Agendado': '#28a745',
    'Interessado - Aguardando confirmação': '#ffc107',
    'Recusou': '#dc3545',
    'Sem Resposta': '#fd7e14',
    'Loja Inativa': '#343a40'
  };

  return (
    <div style={{
      position: 'fixed', top: 0, right: 0, bottom: 0, width: '450px',
      backgroundColor: '#ffffff', boxShadow: '-4px 0 15px rgba(0,0,0,0.15)',
      zIndex: 1000, display: 'flex', flexDirection: 'column', fontFamily: 'Arial, sans-serif'
    }}>
      {/* Cabeçalho */}
      <div style={{ backgroundColor: '#002060', color: '#fff', padding: '16px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <span style={{ fontSize: '12px', background: '#e0a96d', color: '#000', padding: '2px 8px', borderRadius: '4px', fontWeight: 'bold' }}>
            PV: {loja.pvAbadi}
          </span>
          <h3 style={{ margin: '6px 0 0 0', fontSize: '16px', color: '#fff' }}>{loja.razaoSocial}</h3>
        </div>
        <button onClick={onClose} style={{ background: 'none', border: 'none', color: '#fff', fontSize: '20px', cursor: 'pointer' }}>✕</button>
      </div>

      {/* Corpo com Scroll */}
      <div style={{ padding: '20px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
        
        {/* Status Atual */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8f9fa', padding: '10px 12px', borderRadius: '6px' }}>
          <span style={{ fontSize: '13px', fontWeight: 'bold', color: '#555' }}>Status do Registro:</span>
          <span style={{
            backgroundColor: badgeColors[formData.statusContato] || '#6c757d',
            color: '#fff', padding: '4px 10px', borderRadius: '12px', fontSize: '12px', fontWeight: 'bold'
          }}>
            {formData.statusContato}
          </span>
        </div>

        {/* Bloco 1: Contato do Decisor */}
        <div style={{ border: '1px solid #e2e8f0', padding: '12px', borderRadius: '6px' }}>
          <h4 style={{ margin: '0 0 10px 0', fontSize: '13px', color: '#002060', borderBottom: '1px solid #eee', pb: '4px' }}>👤 Dados do Decisor</h4>
          
          <label style={{ fontSize: '12px', color: '#666' }}>Nome do Contato:</label>
          <input type="text" value={formData.nomeContato} onChange={(e) => handleChange('nomeContato', e.target.value)} style={{ width: '100%', padding: '8px', marginBottom: '8px', borderRadius: '4px', border: '1px solid #ccc' }} />

          <div style={{ display: 'flex', gap: '8px' }}>
            <div style={{ flex: 1 }}>
              <label style={{ fontSize: '12px', color: '#666' }}>Telefone / WhatsApp:</label>
              <div style={{ display: 'flex', gap: '4px' }}>
                <input type="text" value={formData.telefoneContato} onChange={(e) => handleChange('telefoneContato', e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} />
                <button type="button" onClick={handleOpenWhatsApp} title="Abrir WhatsApp" style={{ background: '#25D366', border: 'none', color: '#fff', borderRadius: '4px', padding: '0 10px', cursor: 'pointer', fontWeight: 'bold' }}>💬</button>
              </div>
            </div>
          </div>
        </div>

        {/* Bloco 2: Resultado do Atendimento */}
        <div style={{ border: '1px solid #e2e8f0', padding: '12px', borderRadius: '6px' }}>
          <h4 style={{ margin: '0 0 10px 0', fontSize: '13px', color: '#002060', borderBottom: '1px solid #eee', pb: '4px' }}>📞 Resultado da Abordagem</h4>
          
          <label style={{ fontSize: '12px', color: '#666' }}>Status do Contato:</label>
          <select value={formData.statusContato} onChange={(e) => handleChange('statusContato', e.target.value)} style={{ width: '100%', padding: '8px', marginBottom: '8px', borderRadius: '4px', border: '1px solid #ccc' }}>
            <option value="A Contatar">A Contatar</option>
            <option value="Interessado - Aguardando confirmação">Interessado - Aguardando confirmação</option>
            <option value="Agendado">Agendado</option>
            <option value="Recusou">Recusou</option>
            <option value="Sem Resposta">Sem Resposta</option>
            <option value="Loja Inativa">Loja Inativa</option>
          </select>

          {formData.statusContato === 'Agendado' && (
            <div>
              <label style={{ fontSize: '12px', color: '#28a745', fontWeight: 'bold' }}>Data Prevista para o Retreinamento:</label>
              <input type="date" value={formData.dataAgendamento} onChange={(e) => handleChange('dataAgendamento', e.target.value)} style={{ width: '100%', padding: '8px', marginBottom: '8px', borderRadius: '4px', border: '1px solid #28a745' }} />
            </div>
          )}

          {formData.statusContato === 'Recusou' && (
            <div>
              <label style={{ fontSize: '12px', color: '#dc3545', fontWeight: 'bold' }}>Motivo da Recusa:</label>
              <select value={formData.motivoRecusa} onChange={(e) => handleChange('motivoRecusa', e.target.value)} style={{ width: '100%', padding: '8px', marginBottom: '8px', borderRadius: '4px', border: '1px solid #dc3545' }}>
                <option value="">Selecione o motivo...</option>
                <option value="Preço/Custo">Sem orçamento / Custo</option>
                <option value="Sem Tempo/Agenda">Indisponibilidade de agenda</option>
                <option value="Treinamento Recente">Alega ter feito recentemente</option>
                <option value="Outro">Outro motivo</option>
              </select>
            </div>
          )}

          <label style={{ fontSize: '12px', color: '#666' }}>Observações do Atendimento:</label>
          <textarea rows="3" value={formData.observacoes} onChange={(e) => handleChange('observacoes', e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc' }} />
        </div>

        {/* Bloco 3: Mapeamento e Instrutor Sugerido */}
        {instrutoresMaisProximos && instrutoresMaisProximos.length > 0 && (
          <div style={{ border: '1px solid #e2e8f0', padding: '12px', borderRadius: '6px', background: '#f0f4f8' }}>
            <h4 style={{ margin: '0 0 8px 0', fontSize: '13px', color: '#002060' }}>📍 Instrutor Mais Próximo</h4>
            <div style={{ fontSize: '12px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span><strong>{instrutoresMaisProximos[0].nome}</strong> ({instrutoresMaisProximos[0].distanciaKm} km)</span>
              <button 
                type="button" 
                onClick={() => window.open(`https://www.google.com/maps/dir/?api=1&destination=${loja.latitude},${loja.longitude}`, '_blank')}
                style={{ background: '#002060', color: '#fff', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '11px' }}
              >
                Ver Rota 🗺️
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Rodapé fixo */}
      <div style={{ padding: '12px 20px', borderTop: '1px solid #eee', display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#fafafa' }}>
        <span style={{ fontSize: '11px', color: '#888' }}>
          {rascunhoSalvo ? '💾 Salvando...' : '✓ Alterações retidas'}
        </span>
        <button 
          onClick={() => onSave(loja.pvAbadi, formData)}
          style={{ backgroundColor: '#e0a96d', color: '#000', border: 'none', padding: '10px 20px', borderRadius: '4px', fontWeight: 'bold', cursor: 'pointer' }}
        >
          Salvar Registro
        </button>
      </div>
    </div>
  );
};

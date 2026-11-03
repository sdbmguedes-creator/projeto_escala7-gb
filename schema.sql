-- Cole este SQL no SQL Editor do Supabase (https://supabase.com/dashboard/project/.../sql)

CREATE TABLE militares (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  matricula TEXT UNIQUE NOT NULL,
  nome TEXT NOT NULL,
  secao TEXT NOT NULL CHECK (secao IN ('1ªSB','2ªSB','3ªSB','4ªSB','ADM')),
  escala TEXT NOT NULL CHECK (escala IN ('ALFA','BRAVO','CHARLIE','DELTA','ADM','X - A','X - B','X - C','X - D','X - ADM')),
  motorista BOOLEAN DEFAULT FALSE,
  ativo BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

CREATE TABLE grades_diarias (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  competencia DATE NOT NULL,
  dia INTEGER NOT NULL CHECK (dia BETWEEN 1 AND 31),
  militar_id UUID REFERENCES militares(id),
  codigo TEXT NOT NULL CHECK (codigo IN ('X','PJ','PJ12','PJ12H','S1','SPJ','SB','N')),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE(competencia, dia, militar_id)
);

-- Realtime: habilite no painel do Supabase em Database → Realtime → adicione as tabelas militares e grades_diarias.
-- (Não existe comando SQL puro para enable realtime.)

-- Índices para consulta rápida
CREATE INDEX idx_grades_dia ON grades_diarias(competencia, dia);
CREATE INDEX idx_grades_militar ON grades_diarias(militar_id);

-- Exemplo de cadastro (substitua por seus militares reais)
INSERT INTO militares (matricula, nome, secao, escala, motorista) VALUES
('707058-6', 'FULANO DA SILVA', '2ªSB', 'ALFA', true),
('718134-5', 'CICLANO SOUZA', '2ªSB', 'ALFA', false);
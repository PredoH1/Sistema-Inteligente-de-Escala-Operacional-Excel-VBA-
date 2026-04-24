# Sistema Inteligente de Escala Operacional (Excel + VBA)

## 🚀 Visão Geral

Este projeto consiste em um sistema automatizado de geração de escala operacional desenvolvido em **Microsoft Excel utilizando VBA**, com foco em **distribuição equitativa de tarefas**, **redução de esforço manual** e **adaptação dinâmica a eventos operacionais**.

O sistema aplica conceitos de **balanceamento de carga**, **rotação circular (Round-Robin)** e **controle de ciclo**.

---

## 🎯 Problema

Antes da solução:

- Escalas montadas manualmente  
- Distribuição desigual de tarefas  
- Dificuldade em lidar com faltas e feriados  
- Falta de histórico confiável  
- Alto risco de erro operacional  

---

## ⚙️ Solução

Desenvolvimento de um sistema automatizado capaz de:

- Gerar escala semanal automaticamente  
- Balancear carga de trabalho entre colaboradores  
- Adaptar-se a faltas em tempo real  
- Considerar feriados automaticamente  
- Manter histórico mensal e acumulado  
- Garantir rotação justa entre os colaboradores  

---

## 🧠 Lógica do Sistema

### 🔹 Ordenação Multicritério
Os colaboradores são priorizados com base em:

1. Menor carga no mês  
2. Menor carga geral  
3. Ordem alfabética (desempate)  

---

### 🔹 Formação de Duplas
Implementação de:

> **Round-Robin com rotação circular**

- Formação sequencial de pares  
- Reinício automático ao atingir o fim da lista  

---

### 🔹 Controle de Ciclo (Fairness Engine)

- Cada colaborador possui um contador de participação  
- Ao atingir o limite (ex: 2), ocorre:

✔ Reset automático  
✔ Embaralhamento das posições  

---

### 🔹 Embaralhamento

Aplicação de lógica baseada no algoritmo:

> **Fisher-Yates Shuffle**

Objetivo:
- Evitar repetição de duplas  
- Aumentar diversidade  

---

### 🔹 Tratamento de Faltas

O sistema realiza:

- Identificação do colaborador ausente  
- Seleção automática de substituto (menor carga)  
- Reorganização da escala em tempo real  
- Reposição do faltante em dia futuro  

---

### 🔹 Tratamento de Feriados

- Bloqueio automático de datas  
- Preservação da lógica de escala  

---

### 🔹 Persistência de Dados

Controle de:

- Frequência mensal  
- Histórico acumulado  
- Ciclo de participação  

---

## 🏗️ Arquitetura

O sistema é estruturado em múltiplas abas:

| Aba | Função |
|-----|--------|
| ESCALA | Interface principal e execução |
| HISTORICO | Base de dados (mês, geral, ciclo) |
| FERIADOS + FOLGAS | Restrições do sistema |
| INDICADORES | Histórico mensal para análise |

---

## 📸 Demonstração

### 📌 Tela Principal (Escala)
<img width="1251" height="271" alt="image" src="https://github.com/user-attachments/assets/ee2dda4d-0d4c-44b2-a9b7-0b6b6bf226f8" />


### 📌 Base de Dados (Histórico)
<img width="828" height="161" alt="image" src="https://github.com/user-attachments/assets/a023ce52-313a-482a-9502-b9c4963f8eb6" />

---

## 📈 Resultados

- 100% de automação da geração de escala  
- Redução significativa de erros operacionais  
- Distribuição equilibrada de tarefas  
- Maior transparência e controle  
- Facilidade de manutenção e adaptação  

---

## 🛠️ Tecnologias Utilizadas

- Microsoft Excel  
- VBA (Visual Basic for Applications)

---

## 🧩 Conceitos Aplicados

- Load Balancing (Balanceamento de carga)  
- Round-Robin Scheduling  
- Algoritmos de ordenação  
- Randomização (Shuffle)  
- Controle de estado  
- Automação de processos  
- Manipulação de arrays em memória  

---

## 🚀 Como Usar

1. Inserir a data inicial na aba **ESCALA (C1)**  
2. O sistema gera automaticamente a escala semanal  
3. Registrar faltas na coluna correspondente  
4. Finalizar a semana (atualiza contadores)  
5. Fechar o mês para gerar histórico  

---

## 📌 Melhorias Futuras

- Dashboard interativo (Power BI / Excel)  
- Controle de restrições entre colaboradores  
- Sistema de log de auditoria  
- Interface mais amigável (UserForm)  
- Migração para aplicação em Python  

---

## 👨‍💻 Autor

Desenvolvido por [Pedro Henrique Souza Candido]


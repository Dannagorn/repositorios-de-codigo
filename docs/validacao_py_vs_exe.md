# Validação de paridade `.py` vs `.exe`

## Procedimento de regressão
1. Separar conjunto fixo de PDFs de referência.
2. Executar no `.py` e exportar Excel A.
3. Executar no `.exe` e exportar Excel B.
4. Comparar colunas críticas por linha:
   - `tipo_comprovante`
   - `recebedor`
   - `documento_favorecido`
   - `banco`
   - `valor`
   - `data_pagamento`
   - `confianca_extracao`
5. Diferenças aceitas: `estrategia_leitura`, logs e tempos.

## Checklist de build
- [ ] `tesseract.exe` acessível
- [ ] `pdftoppm.exe` acessível
- [ ] `python organizador_comprovantes.py --diagnostico-ambiente`
- [ ] `pyinstaller pyinstaller/organizador_comprovantes.spec`
- [ ] Smoke test em lote real

## Métricas mínimas
- taxa OCR
- tempo total
- tempo médio por arquivo
- total de falhas

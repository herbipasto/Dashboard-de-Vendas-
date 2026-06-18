let dadosGlobais = [];
let mixChart, categoriaMetaChart, evolucaoMensalChart;

// Metas Fixas Totais
const META_FATURAMENTO_TOTAL = 3193569.52;
const META_VOLUME_TOTAL = 1896356;

// Metas por Linha
const METAS_POR_LINHA = {
  "CORTE": 1320000,
  "LEITE": 374000,
  "EQUINOS": 181500,
  "VAREJO": 20856,
  "PREMIUM": 417197
};

// Importar Excel
document.getElementById("btnImportExcel").addEventListener("click", () => {
  document.getElementById("inputExcel").click();
});

document.getElementById('inputExcel').addEventListener('change', function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const vendasRaw = XLSX.utils.sheet_to_json(workbook.Sheets["Vendas"] || {});
    const categoriasRaw = XLSX.utils.sheet_to_json(workbook.Sheets["Categorias"] || {});

    const mapCategorias = {};
    categoriasRaw.forEach(c => {
      const cod = String(c.Código).trim();
      if (!mapCategorias[cod]) mapCategorias[cod] = [];
      const catNome = String(c.Categoria).trim().toUpperCase();
      if(!mapCategorias[cod].includes(catNome)) mapCategorias[cod].push(catNome);
    });

    dadosGlobais = vendasRaw.map(v => {
      const codProd = String(v.Produto).trim();
      const nomeCompleto = String(v["Nome Vendedor"] || "").trim();
      return {
        VendedorCurto: nomeCompleto.split(" ").slice(0, 2).join(" "),
        Faturamento: parseFloat(v.Faturamento) || 0,
        Volume: parseFloat(v.Peso) || 0,
        Categorias: mapCategorias[codProd] || ["OUTROS"]
      };
    });

    atualizarFiltroVendedores();
    processar();
  };
  reader.readAsArrayBuffer(file);
});

// Atualizar filtro de vendedores
function atualizarFiltroVendedores() {
  const select = document.getElementById("filtroVendedor");
  select.innerHTML = '<option value="">Vendedores</option>';
  const vendedoresUnicos = [...new Set(dadosGlobais.map(d => d.VendedorCurto))].sort();
  vendedoresUnicos.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  });
}

document.getElementById("filtroVendedor").addEventListener("change", processar);
document.getElementById("filtroCategoria").addEventListener("change", processar);

// Processar dados e atualizar KPIs
function processar() {
  const vendedorSel = document.getElementById("filtroVendedor").value;
  const categoriaSel = document.getElementById("filtroCategoria").value;

  let filtrados = dadosGlobais.filter(d => {
    const matchVendedor = !vendedorSel || d.VendedorCurto === vendedorSel;
    const matchCategoria = !categoriaSel || d.Categorias.includes(categoriaSel.toUpperCase());
    return matchVendedor && matchCategoria;
  });

  const totalFat = filtrados.reduce((acc, c) => acc + c.Faturamento, 0);
  const totalVol = filtrados.reduce((acc, c) => acc + c.Volume, 0);
  const pFat = (totalFat / META_FATURAMENTO_TOTAL) * 100;
  const pVol = (totalVol / META_VOLUME_TOTAL) * 100;

  document.getElementById("kpi-total").innerText = totalFat.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  document.getElementById("kpi-volume").innerText = totalVol.toLocaleString('pt-BR') + ' ton';
  document.getElementById("kpi-meta-faturamento").innerText = `R$ ${META_FATURAMENTO_TOTAL.toLocaleString('pt-BR')} (${pFat.toFixed(1)}%)`;
  document.getElementById("kpi-meta-volume").innerText = `${META_VOLUME_TOTAL.toLocaleString('pt-BR')} ton (${pVol.toFixed(1)}%)`;

  atualizarMixVendedor(filtrados);
  atualizarCategoriaMeta(filtrados);
  atualizarEvolucaoMensal(filtrados);
}

// Gráfico Mix por Vendedor
function atualizarMixVendedor(dados) {
  if (mixChart) mixChart.destroy();
  const resumo = {};
  dados.forEach(v => {
    resumo[v.VendedorCurto] = (resumo[v.VendedorCurto] || 0) + v.Volume;
  });
  const labels = Object.keys(resumo).sort();
  mixChart = new Chart(document.getElementById('mixChart'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Volume (ton)',
        data: labels.map(l => resumo[l]),
        backgroundColor: '#4caf50'
      }]
    },
    options: {
      plugins: {
        title: { display: true, text: 'Mix por Vendedor (Volume)' },
        datalabels: { anchor: 'end', align: 'top', formatter: v => v.toLocaleString('pt-BR', {maximumFractionDigits:0}) }
      }
    },
    plugins: [ChartDataLabels]
  });
}

// Gráfico Volume por Categoria
function atualizarCategoriaMeta(dados) {
  if (categoriaMetaChart) categoriaMetaChart.destroy();
  const cats = ["CORTE", "LEITE", "EQUINOS", "VAREJO", "PREMIUM"];
  const vends = [...new Set(dados.map(d => d.VendedorCurto))];

  const datasets = vends.map((v, i) => ({
    label: v,
    data: cats.map(c => dados.filter(d => d.Categorias.includes(c) && d.VendedorCurto === v).reduce((a, b) => a + b.Volume, 0)),
    backgroundColor: `hsl(${i * 40}, 70%, 50%)`
  }));

  categoriaMetaChart = new Chart(document.getElementById('categoriaMetaChart'), {
    type: 'bar',
    data: { labels: cats, datasets: datasets },
    options: {
      scales: { x: { stacked: true }, y: { stacked: true } },
      plugins: { title: { display: true, text: 'Volume por Categoria' }, datalabels: { display: false } }
    }
  });
}

// Gráfico Realizado vs Meta
function atualizarEvolucaoMensal(dados) {
  if (evolucaoMensalChart) evolucaoMensalChart.destroy();
  const cats = ["CORTE", "LEITE", "EQUINOS", "VAREJO", "PREMIUM"];
  
  const dataRealizado = cats.map(c => dados.filter(d => d.Categorias.includes(c)).reduce((a, b) => a + b.Volume, 0));
  const dataMetas = cats.map(c => METAS_POR_LINHA[c]);

  evolucaoMensalChart = new Chart(document.getElementById('evolucaoMensalChart'), {
    type: 'bar',
    data: {
      labels: cats,
      datasets: [
        { label: 'Realizado (ton)', data: dataRealizado, backgroundColor: '#2196f3' },
        { label: 'Meta Total (ton)', data: dataMetas, backgroundColor: '#ff9800' }
      ]
    },
    options: {
      plugins: {
        title: { display: true, text: 'Realizado por Linha vs Meta Total' },
        datalabels: { 
          anchor: 'end', 
          align: 'top', 
          formatter: v => v.toLocaleString('pt-BR', {maximumFractionDigits:0}) 
        }
      }
    },
    plugins: [ChartDataLabels]
  });
}

// Exportar PDF
window.onload = function() {
  const { jsPDF } = window.jspdf;
  document.getElementById("btnExportPDF").addEventListener("click", () => {
    html2canvas(document.getElementById("dashboardContent"), { scale: 2 }).then(canvas => {
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF("p", "mm", "a4");
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
      pdf.addImage(imgData, "PNG", 0, 0, pdfWidth, pdfHeight);
      pdf.save("dashboard-herbi-pasto.pdf");
    });
  });
};

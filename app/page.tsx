"use client";

import { ReactNode, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";

type Gasto = {
  id?: number;
  ano?: number;
  mes?: number;
  valor: number;
  categoria: string;
  data: string;
  hora: string;
  observacao: string;
};

type Extra = {
  id?: number;
  ano?: number;
  mes?: number;
  nome: string;
  valor: number;
};

type DadosMes = {
  salario: string;
  observacaoMensal: string;
  extras: Extra[];
  gastos: Gasto[];
};

type PlanejamentoPagamento = {
  diasAtePagamento: number;
  mediaDiariaDisponivel: number;
  mediaSemanalDisponivel: number;
  previsaoSaldoDia25: number;
  dataPagamentoFormatada: string;
};

const CHAVE_DADOS_MENSAIS = "controleGastosMensais";
const DIA_PAGAMENTO = 25;

const categoriasPadrao = [
  "MERCADO",
  "ALIMENTAÇÃO",
  "TRANSPORTE",
  "SAÚDE",
  "LAZER",
  "MORADIA",
  "CONTAS",
  "OUTROS",
];

const meses = [
  { valor: "01", nome: "Janeiro" },
  { valor: "02", nome: "Fevereiro" },
  { valor: "03", nome: "Março" },
  { valor: "04", nome: "Abril" },
  { valor: "05", nome: "Maio" },
  { valor: "06", nome: "Junho" },
  { valor: "07", nome: "Julho" },
  { valor: "08", nome: "Agosto" },
  { valor: "09", nome: "Setembro" },
  { valor: "10", nome: "Outubro" },
  { valor: "11", nome: "Novembro" },
  { valor: "12", nome: "Dezembro" },
];

function obterAnoAtual() {
  return String(new Date().getFullYear());
}

function obterMesAtual() {
  return String(new Date().getMonth() + 1).padStart(2, "0");
}

function criarAnos() {
  const anoAtual = new Date().getFullYear();
  const lista: string[] = [];
  for (let ano = anoAtual - 5; ano <= anoAtual + 5; ano++) {
    lista.push(String(ano));
  }
  return lista;
}

function formatarMoeda(valor: number) {
  return valor.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function nomeDoMes(valorMes: string) {
  return meses.find((m) => m.valor === valorMes)?.nome || valorMes;
}

function formatarDataBR(data: string) {
  if (!data) return "-";
  const partes = data.split("-");
  if (partes.length !== 3) return data;
  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function formatarDataExtensa(data: Date) {
  return data.toLocaleDateString("pt-BR");
}

function obterClasseMensagem(mensagem: string) {
  const texto = mensagem.toLowerCase();

  if (texto.includes("erro")) {
    return "border-red-200 bg-red-50 text-red-700";
  }

  if (
    texto.includes("sucesso") ||
    texto.includes("salvo") ||
    texto.includes("carregado") ||
    texto.includes("conectado") ||
    texto.includes("exportado") ||
    texto.includes("adicionado") ||
    texto.includes("excluído") ||
    texto.includes("apagados")
  ) {
    return "border-emerald-200 bg-emerald-50 text-emerald-700";
  }

  return "border-slate-200 bg-slate-50 text-slate-700";
}

function calcularPercentualUsado(totalDisponivel: number, totalGastos: number) {
  if (totalDisponivel <= 0) return 0;
  const percentual = (totalGastos / totalDisponivel) * 100;
  return Math.min(Math.max(percentual, 0), 100);
}

function classeCategoria(categoria: string) {
  switch (categoria) {
    case "MERCADO":
      return "bg-emerald-100 text-emerald-700";
    case "ALIMENTAÇÃO":
      return "bg-orange-100 text-orange-700";
    case "TRANSPORTE":
      return "bg-sky-100 text-sky-700";
    case "SAÚDE":
      return "bg-rose-100 text-rose-700";
    case "LAZER":
      return "bg-violet-100 text-violet-700";
    case "MORADIA":
      return "bg-amber-100 text-amber-700";
    case "CONTAS":
      return "bg-cyan-100 text-cyan-700";
    default:
      return "bg-slate-100 text-slate-700";
  }
}

function ehCompetenciaAtual(ano: string, mes: string) {
  return ano === obterAnoAtual() && mes === obterMesAtual();
}

function calcularPlanejamentoPagamento(restante: number): PlanejamentoPagamento {
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = hoje.getMonth();

  let dataPagamento = new Date(ano, mes, DIA_PAGAMENTO);

  if (hoje.getDate() > DIA_PAGAMENTO) {
    dataPagamento = new Date(ano, mes + 1, DIA_PAGAMENTO);
  }

  const inicioHoje = new Date(
    hoje.getFullYear(),
    hoje.getMonth(),
    hoje.getDate()
  );

  const diffMs = dataPagamento.getTime() - inicioHoje.getTime();
  const diasAtePagamento = Math.max(
    1,
    Math.ceil(diffMs / (1000 * 60 * 60 * 24))
  );

  const mediaDiariaDisponivel = restante / diasAtePagamento;
  const mediaSemanalDisponivel = mediaDiariaDisponivel * 7;

  return {
    diasAtePagamento,
    mediaDiariaDisponivel,
    mediaSemanalDisponivel,
    previsaoSaldoDia25: restante,
    dataPagamentoFormatada: formatarDataExtensa(dataPagamento),
  };
}

function obterFaixaMediaDiaria(mediaDiaria: number) {
  if (mediaDiaria <= 0) {
    return {
      titulo: "Atenção máxima",
      descricao:
        "O saldo restante já não cobre o período até o próximo pagamento.",
      classeContainer: "border-red-200 bg-red-50 text-red-800",
      classeBadge: "bg-red-600 text-white",
    };
  }

  if (mediaDiaria < 20) {
    return {
      titulo: "Muito baixa",
      descricao:
        "A média diária está bastante apertada. Vale conter gastos não essenciais.",
      classeContainer: "border-red-200 bg-red-50 text-red-800",
      classeBadge: "bg-red-600 text-white",
    };
  }

  if (mediaDiaria < 50) {
    return {
      titulo: "Baixa",
      descricao:
        "A margem diária está curta. É recomendável acompanhar os gastos de perto.",
      classeContainer: "border-amber-200 bg-amber-50 text-amber-900",
      classeBadge: "bg-amber-500 text-white",
    };
  }

  if (mediaDiaria < 100) {
    return {
      titulo: "Moderada",
      descricao:
        "A média diária está razoável, mas ainda pede alguma disciplina.",
      classeContainer: "border-yellow-200 bg-yellow-50 text-yellow-900",
      classeBadge: "bg-yellow-500 text-white",
    };
  }

  return {
    titulo: "Confortável",
    descricao:
      "A média diária está saudável para seguir até o próximo pagamento.",
    classeContainer: "border-emerald-200 bg-emerald-50 text-emerald-900",
    classeBadge: "bg-emerald-600 text-white",
  };
}

export default function Home() {
  const anosDisponiveis = useMemo(() => criarAnos(), []);

  const [anoSelecionado, setAnoSelecionado] = useState(obterAnoAtual());
  const [mesSelecionado, setMesSelecionado] = useState(obterMesAtual());

  const [valor, setValor] = useState("");
  const [categoria, setCategoria] = useState("");
  const [data, setData] = useState("");
  const [hora, setHora] = useState("");
  const [observacao, setObservacao] = useState("");

  const [salario, setSalario] = useState("");
  const [editandoSalario, setEditandoSalario] = useState(true);
  const [observacaoMensal, setObservacaoMensal] = useState("");

  const [extras, setExtras] = useState<Extra[]>([]);
  const [mostrarCampoExtra, setMostrarCampoExtra] = useState(false);
  const [nomeExtra, setNomeExtra] = useState("");
  const [valorExtra, setValorExtra] = useState("");

  const [gastos, setGastos] = useState<Gasto[]>([]);
  const [mensagem, setMensagem] = useState("");

  const chaveMesAtual = `${anoSelecionado}-${mesSelecionado}`;

  function lerBancoCompleto(): Record<string, DadosMes> {
    try {
      const bruto = localStorage.getItem(CHAVE_DADOS_MENSAIS);
      return bruto ? JSON.parse(bruto) : {};
    } catch (error) {
      console.error("Erro ao ler base mensal:", error);
      return {};
    }
  }

  function salvarBancoCompleto(banco: Record<string, DadosMes>) {
    localStorage.setItem(CHAVE_DADOS_MENSAIS, JSON.stringify(banco));
  }

  function salvarMesAtualLocal(
    novoSalario: string,
    novaObservacaoMensal: string,
    novosExtras: Extra[],
    novosGastos: Gasto[]
  ) {
    try {
      const banco = lerBancoCompleto();

      banco[chaveMesAtual] = {
        salario: novoSalario,
        observacaoMensal: novaObservacaoMensal,
        extras: novosExtras,
        gastos: novosGastos,
      };

      salvarBancoCompleto(banco);
    } catch (error) {
      console.error("Erro ao salvar dados locais do mês:", error);
    }
  }

  async function testarSupabase() {
    const { error } = await supabase.from("meses").select("*");

    if (error) {
      console.error(error);
      setMensagem("Erro ao conectar com Supabase");
      return;
    }

    setMensagem("Supabase conectado com sucesso!");
  }

  async function carregarSalarioSupabase(ano: string, mes: string) {
    const { data, error } = await supabase
      .from("meses")
      .select("*")
      .eq("ano", Number(ano))
      .eq("mes", Number(mes))
      .maybeSingle();

    if (error) {
      console.error("Erro ao carregar salário do Supabase:", error);
      setMensagem("Erro ao carregar salário do Supabase");
      return null;
    }

    return data;
  }

  async function salvarSalarioSupabase(
    ano: string,
    mes: string,
    salarioAtual: string,
    observacaoMensalAtual: string
  ) {
    const { error } = await supabase.from("meses").upsert(
      [
        {
          ano: Number(ano),
          mes: Number(mes),
          salario: Number(salarioAtual || 0),
          observacao_mensal: observacaoMensalAtual || "",
        },
      ],
      {
        onConflict: "ano,mes",
      }
    );

    if (error) {
      console.error("Erro ao salvar salário no Supabase:", error);
      setMensagem("Erro ao salvar salário no Supabase");
      return false;
    }

    return true;
  }

  async function carregarGastosSupabase(ano: string, mes: string) {
    const { data, error } = await supabase
      .from("gastos")
      .select("*")
      .eq("ano", Number(ano))
      .eq("mes", Number(mes))
      .order("data", { ascending: true })
      .order("hora", { ascending: true });

    if (error) {
      console.error("Erro ao carregar gastos do Supabase:", error);
      setMensagem("Erro ao carregar gastos do Supabase");
      return [];
    }

    return (data || []).map((item) => ({
      id: item.id,
      ano: item.ano,
      mes: item.mes,
      valor: Number(item.valor || 0),
      categoria: item.categoria || "",
      data: item.data || "",
      hora: item.hora || "",
      observacao: item.observacao || "",
    }));
  }

  async function salvarGastoSupabase(gasto: Gasto) {
    const { error } = await supabase.from("gastos").insert([
      {
        ano: Number(anoSelecionado),
        mes: Number(mesSelecionado),
        valor: Number(gasto.valor || 0),
        categoria: gasto.categoria,
        data: gasto.data,
        hora: gasto.hora || "",
        observacao: gasto.observacao || "",
      },
    ]);

    if (error) {
      console.error("Erro ao salvar gasto no Supabase:", error);
      setMensagem("Erro ao salvar gasto no Supabase");
      return false;
    }

    return true;
  }

  async function excluirGastoSupabase(id: number) {
    const { error } = await supabase.from("gastos").delete().eq("id", id);

    if (error) {
      console.error("Erro ao excluir gasto no Supabase:", error);
      setMensagem("Erro ao excluir gasto no Supabase");
      return false;
    }

    return true;
  }

  async function carregarExtrasSupabase(ano: string, mes: string) {
    const { data, error } = await supabase
      .from("extras")
      .select("*")
      .eq("ano", Number(ano))
      .eq("mes", Number(mes))
      .order("id", { ascending: true });

    if (error) {
      console.error("Erro ao carregar extras do Supabase:", error);
      setMensagem("Erro ao carregar extras do Supabase");
      return [];
    }

    return (data || []).map((item) => ({
      id: item.id,
      ano: item.ano,
      mes: item.mes,
      nome: item.nome || "",
      valor: Number(item.valor || 0),
    }));
  }

  async function salvarExtraSupabase(extra: Extra) {
    const { error } = await supabase.from("extras").insert([
      {
        ano: Number(anoSelecionado),
        mes: Number(mesSelecionado),
        nome: extra.nome,
        valor: Number(extra.valor || 0),
      },
    ]);

    if (error) {
      console.error("Erro ao salvar extra no Supabase:", error);
      setMensagem("Erro ao salvar extra no Supabase");
      return false;
    }

    return true;
  }

  async function excluirExtraSupabase(id: number) {
    const { error } = await supabase.from("extras").delete().eq("id", id);

    if (error) {
      console.error("Erro ao excluir extra no Supabase:", error);
      setMensagem("Erro ao excluir extra no Supabase");
      return false;
    }

    return true;
  }

  async function carregarMesAtual() {
    try {
      const banco = lerBancoCompleto();
      const dadosLocais = banco[chaveMesAtual];

      setObservacaoMensal(dadosLocais?.observacaoMensal || "");
      setEditandoSalario(true);

      const dadosSupabase = await carregarSalarioSupabase(
        anoSelecionado,
        mesSelecionado
      );

      if (dadosSupabase) {
        const salarioDoBanco = String(dadosSupabase.salario ?? "");
        const observacaoDoBanco = String(dadosSupabase.observacao_mensal ?? "");

        setSalario(salarioDoBanco);
        setObservacaoMensal(observacaoDoBanco);
        setEditandoSalario(!(salarioDoBanco && salarioDoBanco !== ""));
      } else {
        setSalario(dadosLocais?.salario || "");
        setObservacaoMensal(dadosLocais?.observacaoMensal || "");
        setEditandoSalario(
          !(dadosLocais?.salario && dadosLocais.salario !== "")
        );
      }

      const gastosBanco = await carregarGastosSupabase(
        anoSelecionado,
        mesSelecionado
      );

      if (gastosBanco.length > 0) {
        setGastos(gastosBanco);
      } else {
        setGastos(dadosLocais?.gastos || []);
      }

      const extrasBanco = await carregarExtrasSupabase(
        anoSelecionado,
        mesSelecionado
      );

      if (extrasBanco.length > 0) {
        setExtras(extrasBanco);
      } else {
        setExtras(dadosLocais?.extras || []);
      }

      salvarMesAtualLocal(
        dadosSupabase
          ? String(dadosSupabase.salario ?? "")
          : dadosLocais?.salario || "",
        dadosSupabase
          ? String(dadosSupabase.observacao_mensal ?? "")
          : dadosLocais?.observacaoMensal || "",
        extrasBanco.length > 0 ? extrasBanco : dadosLocais?.extras || [],
        gastosBanco.length > 0 ? gastosBanco : dadosLocais?.gastos || []
      );

      if (
        dadosSupabase ||
        gastosBanco.length > 0 ||
        extrasBanco.length > 0 ||
        dadosLocais
      ) {
        setMensagem(`Dados carregados para ${mesSelecionado}/${anoSelecionado}.`);
      } else {
        setMensagem(
          `Nenhum dado encontrado para ${mesSelecionado}/${anoSelecionado}.`
        );
      }

      setValor("");
      setCategoria("");
      setData("");
      setHora("");
      setObservacao("");
      setMostrarCampoExtra(false);
      setNomeExtra("");
      setValorExtra("");
    } catch (error) {
      console.error("Erro ao carregar mês atual:", error);
      setMensagem("Erro ao carregar os dados do mês.");
    }
  }

  useEffect(() => {
    carregarMesAtual();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [anoSelecionado, mesSelecionado]);

  const adicionarGasto = async () => {
    if (!valor || !categoria || !data) {
      setMensagem("Preencha data, categoria e valor do gasto.");
      return;
    }

    const novoGasto: Gasto = {
      valor: Number(valor),
      categoria,
      data,
      hora,
      observacao,
    };

    const ok = await salvarGastoSupabase(novoGasto);

    if (!ok) {
      setMensagem("Erro ao salvar gasto no banco.");
      return;
    }

    const novosGastosLocais = [...gastos, novoGasto];
    salvarMesAtualLocal(salario, observacaoMensal, extras, novosGastosLocais);

    await carregarMesAtual();

    setValor("");
    setCategoria("");
    setData("");
    setHora("");
    setObservacao("");
    setMensagem("Gasto salvo no Supabase.");
  };

  const excluirGasto = async (indice: number) => {
    const gasto = gastos[indice];

    if (gasto?.id) {
      const ok = await excluirGastoSupabase(gasto.id);
      if (!ok) return;
    } else {
      const novosGastos = gastos.filter((_, i) => i !== indice);
      salvarMesAtualLocal(salario, observacaoMensal, extras, novosGastos);
    }

    await carregarMesAtual();
    setMensagem("Gasto excluído com sucesso.");
  };

  const alterarSalario = (novoValor: string) => {
    setSalario(novoValor);
  };

  const salvarSalario = async () => {
    const ok = await salvarSalarioSupabase(
      anoSelecionado,
      mesSelecionado,
      salario,
      observacaoMensal
    );

    if (!ok) return;

    salvarMesAtualLocal(salario, observacaoMensal, extras, gastos);
    setEditandoSalario(false);
    setMensagem(`Salário salvo para ${mesSelecionado}/${anoSelecionado}.`);
  };

  const editarSalario = () => {
    setEditandoSalario(true);
    setMensagem("Edição do salário liberada.");
  };

  const alterarObservacaoMensal = (novoTexto: string) => {
    setObservacaoMensal(novoTexto);
    salvarMesAtualLocal(salario, novoTexto, extras, gastos);
  };

  const adicionarExtra = async () => {
    if (!nomeExtra || !valorExtra) {
      setMensagem("Preencha o nome e o valor do extra.");
      return;
    }

    const novoExtra: Extra = {
      nome: nomeExtra,
      valor: Number(valorExtra),
    };

    const ok = await salvarExtraSupabase(novoExtra);

    if (!ok) {
      setMensagem("Erro ao salvar extra no banco.");
      return;
    }

    const novosExtrasLocais = [...extras, novoExtra];
    salvarMesAtualLocal(salario, observacaoMensal, novosExtrasLocais, gastos);

    setNomeExtra("");
    setValorExtra("");
    setMostrarCampoExtra(false);

    await carregarMesAtual();
    setMensagem("Extra adicionado com sucesso.");
  };

  const excluirExtra = async (indice: number) => {
    const extra = extras[indice];

    if (extra?.id) {
      const ok = await excluirExtraSupabase(extra.id);
      if (!ok) return;
    } else {
      const novosExtras = extras.filter((_, i) => i !== indice);
      salvarMesAtualLocal(salario, observacaoMensal, novosExtras, gastos);
    }

    await carregarMesAtual();
    setMensagem("Extra excluído com sucesso.");
  };

  const limparMesAtual = async () => {
    try {
      const banco = lerBancoCompleto();
      delete banco[chaveMesAtual];
      salvarBancoCompleto(banco);

      await supabase
        .from("meses")
        .delete()
        .eq("ano", Number(anoSelecionado))
        .eq("mes", Number(mesSelecionado));

      await supabase
        .from("gastos")
        .delete()
        .eq("ano", Number(anoSelecionado))
        .eq("mes", Number(mesSelecionado));

      await supabase
        .from("extras")
        .delete()
        .eq("ano", Number(anoSelecionado))
        .eq("mes", Number(mesSelecionado));

      setSalario("");
      setObservacaoMensal("");
      setExtras([]);
      setGastos([]);
      setValor("");
      setCategoria("");
      setData("");
      setHora("");
      setObservacao("");
      setEditandoSalario(true);
      setMostrarCampoExtra(false);
      setNomeExtra("");
      setValorExtra("");

      setMensagem(`Dados apagados de ${mesSelecionado}/${anoSelecionado}.`);
    } catch (error) {
      console.error("Erro ao limpar mês atual:", error);
      setMensagem("Erro ao limpar os dados do mês.");
    }
  };

  const exportarExcelMesAtual = () => {
    try {
      const nomeMesAtual = nomeDoMes(mesSelecionado);

      const totalExtras = extras.reduce((acc, extra) => acc + extra.valor, 0);
      const totalGastos = gastos.reduce((acc, gasto) => acc + gasto.valor, 0);
      const totalDisponivel = Number(salario || 0) + totalExtras;
      const restante = totalDisponivel - totalGastos;
      const competenciaAtual = ehCompetenciaAtual(anoSelecionado, mesSelecionado);
      const planejamento = calcularPlanejamentoPagamento(restante);

      const abaResumo = [
        { Campo: "Ano", Valor: anoSelecionado },
        { Campo: "Mês", Valor: nomeMesAtual },
        { Campo: "Salário do mês", Valor: Number(salario || 0) },
        { Campo: "Total de extras", Valor: totalExtras },
        { Campo: "Total disponível", Valor: totalDisponivel },
        { Campo: "Total de gastos", Valor: totalGastos },
        { Campo: "Valor restante", Valor: restante },
        { Campo: "Observação do mês", Valor: observacaoMensal || "" },
        {
          Campo: "Dias até o pagamento",
          Valor: competenciaAtual ? planejamento.diasAtePagamento : "N/A",
        },
        {
          Campo: "Média disponível por dia",
          Valor: competenciaAtual ? planejamento.mediaDiariaDisponivel : "N/A",
        },
        {
          Campo: "Média disponível por semana",
          Valor: competenciaAtual ? planejamento.mediaSemanalDisponivel : "N/A",
        },
        {
          Campo: "Previsão de saldo no dia 25",
          Valor: competenciaAtual ? planejamento.previsaoSaldoDia25 : "N/A",
        },
      ];

      const abaExtras =
        extras.length > 0
          ? extras.map((extra) => ({
              Nome: extra.nome,
              Valor: extra.valor,
            }))
          : [{ Nome: "Sem extras", Valor: 0 }];

      const abaGastos =
        gastos.length > 0
          ? gastos.map((gasto) => ({
              Data: gasto.data,
              Hora: gasto.hora || "",
              Categoria: gasto.categoria,
              Observação: gasto.observacao || "",
              Valor: gasto.valor,
            }))
          : [
              {
                Data: "",
                Hora: "",
                Categoria: "Sem gastos",
                Observação: "",
                Valor: 0,
              },
            ];

      const workbook = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.json_to_sheet(abaResumo),
        "Resumo"
      );
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.json_to_sheet(abaExtras),
        "Extras"
      );
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.json_to_sheet(abaGastos),
        "Gastos"
      );

      const nomeArquivo = `controle-gastos-${anoSelecionado}-${mesSelecionado}.xlsx`;
      XLSX.writeFile(workbook, nomeArquivo);

      setMensagem(`Excel exportado: ${nomeArquivo}`);
    } catch (error) {
      console.error("Erro ao exportar Excel do mês:", error);
      setMensagem("Erro ao exportar o Excel do mês.");
    }
  };

  const totalExtras = extras.reduce((acc, extra) => acc + extra.valor, 0);
  const totalDisponivel = Number(salario || 0) + totalExtras;
  const totalGastos = gastos.reduce((acc, gasto) => acc + gasto.valor, 0);
  const restante = totalDisponivel - totalGastos;
  const percentualUsado = calcularPercentualUsado(totalDisponivel, totalGastos);
  const nomeMesAtual = nomeDoMes(mesSelecionado);
  const competenciaAtual = ehCompetenciaAtual(anoSelecionado, mesSelecionado);
  const planejamentoPagamento = calcularPlanejamentoPagamento(restante);
  const faixaMediaDiaria = obterFaixaMediaDiaria(
    planejamentoPagamento.mediaDiariaDisponivel
  );

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_top,_#f8fafc,_#e2e8f0_55%,_#cbd5e1)] text-slate-900">
      <div className="mx-auto max-w-7xl px-4 py-6 sm:px-6 lg:px-8">
        <section className="overflow-hidden rounded-[32px] border border-slate-200 bg-white shadow-[0_20px_60px_rgba(15,23,42,0.10)]">
          <div className="bg-gradient-to-r from-slate-950 via-slate-800 to-slate-700 px-6 py-8 text-white sm:px-8">
            <div className="flex flex-col gap-6 xl:flex-row xl:items-center xl:justify-between">
              <div className="max-w-3xl">
                <div className="mb-3 inline-flex rounded-full border border-white/15 bg-white/10 px-4 py-1.5 text-xs font-semibold uppercase tracking-[0.22em] text-slate-100">
                  Sistema financeiro pessoal
                </div>

                <h1 className="text-3xl font-bold tracking-tight sm:text-4xl">
                  Controle de Gastos
                </h1>

                <p className="mt-3 max-w-2xl text-sm leading-6 text-slate-200 sm:text-base">
                  Organize salário, extras, despesas e observações mensais com
                  separação por ano e mês, mantendo integração com Supabase,
                  exportação em Excel e indicadores de planejamento até o dia 25.
                </p>
              </div>

              <div className="grid gap-3 sm:grid-cols-2 xl:min-w-[380px]">
                <HeroMiniCard
                  titulo="Competência"
                  valor={`${nomeMesAtual}/${anoSelecionado}`}
                />
                <HeroMiniCard
                  titulo="Saldo restante"
                  valor={formatarMoeda(restante)}
                />
                <HeroMiniCard
                  titulo="Qtd. de gastos"
                  valor={String(gastos.length)}
                />
                <HeroMiniCard
                  titulo="Qtd. de extras"
                  valor={String(extras.length)}
                />
              </div>
            </div>
          </div>

          <div className="border-t border-slate-200 bg-white px-6 py-4 sm:px-8">
            <div
              className={`rounded-2xl border px-4 py-4 text-sm ${obterClasseMensagem(
                mensagem
              )}`}
            >
              <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
                <div>
                  <span className="font-semibold">Status:</span>{" "}
                  {mensagem || "Sistema pronto para uso."}
                </div>

                <div className="flex flex-wrap gap-3">
                  <button
                    onClick={testarSupabase}
                    className="rounded-2xl bg-slate-900 px-4 py-2.5 font-medium text-white transition hover:bg-slate-800"
                  >
                    Testar Supabase
                  </button>

                  <button
                    onClick={carregarMesAtual}
                    className="rounded-2xl border border-slate-300 bg-white px-4 py-2.5 font-medium text-slate-700 transition hover:bg-slate-50"
                  >
                    Recarregar mês
                  </button>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <IndicadorResumo
            titulo="Salário do mês"
            valor={formatarMoeda(Number(salario || 0))}
            descricao="Base mensal registrada"
            cor="blue"
          />
          <IndicadorResumo
            titulo="Total de extras"
            valor={formatarMoeda(totalExtras)}
            descricao="Entradas adicionais"
            cor="green"
          />
          <IndicadorResumo
            titulo="Total gasto"
            valor={formatarMoeda(totalGastos)}
            descricao="Somatório das despesas"
            cor="red"
          />
          <IndicadorResumo
            titulo="Total disponível"
            valor={formatarMoeda(totalDisponivel)}
            descricao="Salário + extras"
            cor="violet"
          />
        </section>

        <section className="mt-6 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_14px_40px_rgba(15,23,42,0.08)]">
          <div className="flex flex-col gap-5 xl:flex-row xl:items-end xl:justify-between">
            <div>
              <h2 className="text-xl font-bold text-slate-900">Competência</h2>
              <p className="mt-1 text-sm text-slate-500">
                Escolha o período que deseja consultar ou editar.
              </p>
            </div>

            <div className="grid w-full gap-4 md:grid-cols-2 xl:w-[520px]">
              <CampoBox label="Ano">
                <select
                  value={anoSelecionado}
                  onChange={(e) => setAnoSelecionado(e.target.value)}
                  className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                >
                  {anosDisponiveis.map((ano) => (
                    <option key={ano} value={ano}>
                      {ano}
                    </option>
                  ))}
                </select>
              </CampoBox>

              <CampoBox label="Mês">
                <select
                  value={mesSelecionado}
                  onChange={(e) => setMesSelecionado(e.target.value)}
                  className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                >
                  {meses.map((mes) => (
                    <option key={mes.valor} value={mes.valor}>
                      {mes.nome}
                    </option>
                  ))}
                </select>
              </CampoBox>
            </div>
          </div>

          <div className="mt-6 grid gap-4 xl:grid-cols-[1.2fr_0.8fr]">
            <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <p className="text-sm font-semibold text-slate-700">
                    Consumo do orçamento mensal
                  </p>
                  <p className="text-sm text-slate-500">
                    Relação entre gastos e total disponível
                  </p>
                </div>

                <div className="rounded-full bg-white px-4 py-2 text-sm font-bold text-slate-700 shadow-sm">
                  {percentualUsado.toFixed(1)}%
                </div>
              </div>

              <div className="mt-4 h-4 overflow-hidden rounded-full bg-slate-200">
                <div
                  className="h-full rounded-full bg-gradient-to-r from-blue-600 via-cyan-500 to-emerald-500 transition-all duration-300"
                  style={{ width: `${percentualUsado}%` }}
                />
              </div>

              <div className="mt-4 grid gap-3 sm:grid-cols-3">
                <MiniResumo
                  titulo="Disponível"
                  valor={formatarMoeda(totalDisponivel)}
                />
                <MiniResumo
                  titulo="Gasto"
                  valor={formatarMoeda(totalGastos)}
                />
                <MiniResumo
                  titulo="Restante"
                  valor={formatarMoeda(restante)}
                />
              </div>
            </div>

            <div className="rounded-3xl border border-slate-200 bg-gradient-to-br from-slate-900 via-slate-800 to-slate-700 p-5 text-white">
              <p className="text-sm font-semibold uppercase tracking-wide text-slate-200">
                Ações rápidas
              </p>

              <div className="mt-4 grid gap-3">
                <button
                  onClick={exportarExcelMesAtual}
                  className="rounded-2xl bg-emerald-500 px-4 py-3 font-medium text-white transition hover:bg-emerald-600"
                >
                  Exportar Excel do mês
                </button>

                <button
                  onClick={() => setMostrarCampoExtra(!mostrarCampoExtra)}
                  className="rounded-2xl bg-white/10 px-4 py-3 font-medium text-white transition hover:bg-white/15"
                >
                  {mostrarCampoExtra ? "Cancelar extra" : "Adicionar extra"}
                </button>

                <button
                  onClick={limparMesAtual}
                  className="rounded-2xl bg-red-600 px-4 py-3 font-medium text-white transition hover:bg-red-700"
                >
                  Limpar mês
                </button>
              </div>
            </div>
          </div>
        </section>

        <section className="mt-6 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_14px_40px_rgba(15,23,42,0.08)]">
          <div className="flex flex-col gap-2 border-b border-slate-200 pb-5">
            <h2 className="text-xl font-bold text-slate-900">
              Planejamento até o pagamento
            </h2>
            <p className="text-sm text-slate-500">
              Indicadores baseados no saldo restante e no próximo recebimento do
              dia 25.
            </p>
          </div>

          {competenciaAtual ? (
            <>
              <div className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                <IndicadorResumo
                  titulo="Dias até pagamento"
                  valor={String(planejamentoPagamento.diasAtePagamento)}
                  descricao={`Próximo pagamento em ${planejamentoPagamento.dataPagamentoFormatada}`}
                  cor="blue"
                />
                <IndicadorResumo
                  titulo="Disponível por dia"
                  valor={formatarMoeda(
                    planejamentoPagamento.mediaDiariaDisponivel
                  )}
                  descricao="Média diária até o dia 25"
                  cor="green"
                />
                <IndicadorResumo
                  titulo="Disponível por semana"
                  valor={formatarMoeda(
                    planejamentoPagamento.mediaSemanalDisponivel
                  )}
                  descricao="Média semanal projetada"
                  cor="violet"
                />
                <IndicadorResumo
                  titulo="Previsão no dia 25"
                  valor={formatarMoeda(
                    planejamentoPagamento.previsaoSaldoDia25
                  )}
                  descricao="Saldo estimado mantendo o cenário atual"
                  cor="red"
                />
              </div>

              <div
                className={`mt-5 rounded-3xl border p-5 ${faixaMediaDiaria.classeContainer}`}
              >
                <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
                  <div>
                    <p className="text-sm font-semibold uppercase tracking-wide">
                      Alerta de média diária
                    </p>
                    <h3 className="mt-1 text-xl font-bold">
                      Situação: {faixaMediaDiaria.titulo}
                    </h3>
                    <p className="mt-2 text-sm leading-6">
                      {faixaMediaDiaria.descricao}
                    </p>
                  </div>

                  <div
                    className={`w-fit rounded-full px-4 py-2 text-sm font-semibold ${faixaMediaDiaria.classeBadge}`}
                  >
                    Média diária:{" "}
                    {formatarMoeda(planejamentoPagamento.mediaDiariaDisponivel)}
                  </div>
                </div>
              </div>

              <div className="mt-5 grid gap-4 xl:grid-cols-3">
                <MiniResumo
                  titulo="Saldo restante atual"
                  valor={formatarMoeda(restante)}
                />
                <MiniResumo
                  titulo="Limite médio por dia"
                  valor={formatarMoeda(
                    planejamentoPagamento.mediaDiariaDisponivel
                  )}
                />
                <MiniResumo
                  titulo="Limite médio por semana"
                  valor={formatarMoeda(
                    planejamentoPagamento.mediaSemanalDisponivel
                  )}
                />
              </div>
            </>
          ) : (
            <div className="mt-6 rounded-3xl border border-dashed border-slate-300 bg-slate-50 px-6 py-8">
              <p className="text-base font-semibold text-slate-800">
                Indicadores de planejamento indisponíveis para esta competência
              </p>
              <p className="mt-2 text-sm leading-6 text-slate-600">
                Os cálculos de média diária, média semanal, dias até pagamento e
                previsão de saldo no dia 25 aparecem apenas quando o mês e o ano
                selecionados correspondem à competência atual.
              </p>
            </div>
          )}
        </section>

        <div className="mt-6 grid gap-6 xl:grid-cols-[1.08fr_0.92fr]">
          <section className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_14px_40px_rgba(15,23,42,0.08)]">
            <div className="flex flex-col gap-2 border-b border-slate-200 pb-5">
              <h2 className="text-xl font-bold text-slate-900">Dados do mês</h2>
              <p className="text-sm text-slate-500">
                Área para salário, observações mensais e entradas extras.
              </p>
            </div>

            <div className="mt-6 rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h3 className="text-lg font-semibold text-slate-900">
                    Salário do mês
                  </h3>
                  <p className="text-sm text-slate-500">
                    Valor base usado no cálculo do período.
                  </p>
                </div>

                <span
                  className={`w-fit rounded-full px-3 py-1 text-xs font-semibold ${
                    editandoSalario
                      ? "bg-amber-100 text-amber-700"
                      : "bg-emerald-100 text-emerald-700"
                  }`}
                >
                  {editandoSalario ? "Editando" : "Salvo"}
                </span>
              </div>

              <div className="flex flex-col gap-3 lg:flex-row">
                <input
                  type="number"
                  placeholder="Digite o salário do mês"
                  value={salario}
                  disabled={!editandoSalario}
                  onChange={(e) => alterarSalario(e.target.value)}
                  className={`w-full rounded-2xl border px-4 py-3 outline-none transition ${
                    editandoSalario
                      ? "border-slate-300 bg-white focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                      : "border-slate-200 bg-slate-100 text-slate-700"
                  }`}
                />

                {editandoSalario ? (
                  <button
                    onClick={salvarSalario}
                    className="rounded-2xl bg-blue-600 px-5 py-3 font-medium text-white transition hover:bg-blue-700"
                  >
                    Salvar salário
                  </button>
                ) : (
                  <button
                    onClick={editarSalario}
                    className="rounded-2xl bg-slate-700 px-5 py-3 font-medium text-white transition hover:bg-slate-800"
                  >
                    Editar salário
                  </button>
                )}
              </div>
            </div>

            <div className="mt-6 rounded-3xl border border-slate-200 bg-white p-5">
              <h3 className="text-lg font-semibold text-slate-900">
                Observação do mês
              </h3>
              <p className="mt-1 text-sm text-slate-500">
                Use este campo para anotações gerais da competência.
              </p>

              <textarea
                value={observacaoMensal}
                onChange={(e) => alterarObservacaoMensal(e.target.value)}
                rows={4}
                placeholder="Digite observações gerais deste mês..."
                className="mt-4 w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
              />
            </div>

            {mostrarCampoExtra && (
              <div className="mt-6 rounded-3xl border border-emerald-200 bg-emerald-50 p-5">
                <h3 className="text-lg font-semibold text-emerald-900">
                  Novo extra
                </h3>
                <p className="mt-1 text-sm text-emerald-700">
                  Adicione um valor que deve somar ao total disponível do mês.
                </p>

                <div className="mt-4 grid gap-3 md:grid-cols-2">
                  <input
                    type="text"
                    placeholder="Nome do extra"
                    value={nomeExtra}
                    onChange={(e) => setNomeExtra(e.target.value)}
                    className="w-full rounded-2xl border border-emerald-200 bg-white px-4 py-3 outline-none transition focus:border-emerald-400 focus:ring-2 focus:ring-emerald-100"
                  />

                  <input
                    type="number"
                    placeholder="Valor do extra"
                    value={valorExtra}
                    onChange={(e) => setValorExtra(e.target.value)}
                    className="w-full rounded-2xl border border-emerald-200 bg-white px-4 py-3 outline-none transition focus:border-emerald-400 focus:ring-2 focus:ring-emerald-100"
                  />
                </div>

                <div className="mt-4">
                  <button
                    onClick={adicionarExtra}
                    className="rounded-2xl bg-emerald-600 px-5 py-3 font-medium text-white transition hover:bg-emerald-700"
                  >
                    Salvar extra
                  </button>
                </div>
              </div>
            )}

            <div className="mt-6 rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h3 className="text-lg font-semibold text-slate-900">
                    Extras adicionados
                  </h3>
                  <p className="text-sm text-slate-500">
                    Valores complementares cadastrados no mês.
                  </p>
                </div>

                <div className="rounded-full bg-white px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm">
                  {extras.length} item(ns)
                </div>
              </div>

              {extras.length === 0 ? (
                <div className="mt-5">
                  <EstadoVazio
                    titulo="Nenhum extra cadastrado"
                    descricao="Quando você adicionar um extra, ele aparecerá aqui."
                  />
                </div>
              ) : (
                <div className="mt-5 overflow-x-auto rounded-2xl border border-slate-200 bg-white">
                  <table className="min-w-full text-sm">
                    <thead className="bg-slate-100 text-slate-700">
                      <tr>
                        <th className="px-4 py-3 text-left font-semibold">
                          Nome
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Valor
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Ação
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {extras.map((extra, i) => (
                        <tr
                          key={extra.id ?? `${extra.nome}-${extra.valor}-${i}`}
                          className="border-t border-slate-200"
                        >
                          <td className="px-4 py-3 font-medium text-slate-800">
                            {extra.nome}
                          </td>
                          <td className="px-4 py-3 font-semibold text-emerald-700">
                            {formatarMoeda(extra.valor)}
                          </td>
                          <td className="px-4 py-3">
                            <button
                              onClick={() => excluirExtra(i)}
                              className="rounded-xl bg-red-600 px-3 py-2 text-white transition hover:bg-red-700"
                            >
                              Excluir
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <div className="mt-6 rounded-3xl bg-gradient-to-br from-slate-950 via-slate-800 to-slate-700 p-6 text-white shadow-lg">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
                <div>
                  <h3 className="text-xl font-bold">Resumo financeiro</h3>
                  <p className="mt-1 text-sm text-slate-200">
                    Consolidação dos valores da competência atual.
                  </p>
                </div>

                <button
                  onClick={exportarExcelMesAtual}
                  className="rounded-2xl bg-emerald-500 px-5 py-3 font-medium text-white transition hover:bg-emerald-600"
                >
                  Exportar Excel
                </button>
              </div>

              <div className="mt-5 grid gap-4 md:grid-cols-2">
                <ResumoEscuro
                  titulo="Salário"
                  valor={formatarMoeda(Number(salario || 0))}
                />
                <ResumoEscuro
                  titulo="Extras"
                  valor={formatarMoeda(totalExtras)}
                />
                <ResumoEscuro
                  titulo="Disponível"
                  valor={formatarMoeda(totalDisponivel)}
                />
                <ResumoEscuro
                  titulo="Gastos"
                  valor={formatarMoeda(totalGastos)}
                />
              </div>

              <div className="mt-5 grid gap-4 md:grid-cols-2">
                <ResumoEscuro
                  titulo="Restante"
                  valor={formatarMoeda(restante)}
                />
                <ResumoEscuro
                  titulo="Previsão no dia 25"
                  valor={
                    competenciaAtual
                      ? formatarMoeda(planejamentoPagamento.previsaoSaldoDia25)
                      : "Somente no mês atual"
                  }
                />
              </div>
            </div>
          </section>

          <section className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_14px_40px_rgba(15,23,42,0.08)]">
            <div className="flex flex-col gap-2 border-b border-slate-200 pb-5">
              <h2 className="text-xl font-bold text-slate-900">Novo gasto</h2>
              <p className="text-sm text-slate-500">
                Cadastre uma nova despesa do mês selecionado.
              </p>
            </div>

            <div className="mt-6 grid gap-4">
              <div className="grid gap-4 md:grid-cols-2">
                <CampoBox label="Data">
                  <input
                    type="date"
                    value={data}
                    onChange={(e) => setData(e.target.value)}
                    className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                  />
                </CampoBox>

                <CampoBox label="Hora">
                  <input
                    type="time"
                    value={hora}
                    onChange={(e) => setHora(e.target.value)}
                    className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                  />
                </CampoBox>
              </div>

              <CampoBox label="Categoria">
                <select
                  value={categoria}
                  onChange={(e) => setCategoria(e.target.value)}
                  className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                >
                  <option value="">Selecione a categoria</option>
                  {categoriasPadrao.map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </CampoBox>

              <CampoBox label="Observação">
                <input
                  placeholder="Ex.: combustível, remédio, mercado da semana..."
                  value={observacao}
                  onChange={(e) => setObservacao(e.target.value)}
                  className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                />
              </CampoBox>

              <CampoBox label="Valor do gasto">
                <input
                  type="number"
                  placeholder="Digite o valor do gasto"
                  value={valor}
                  onChange={(e) => setValor(e.target.value)}
                  className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
                />
              </CampoBox>
            </div>

            <div className="mt-6 flex flex-wrap gap-3">
              <button
                onClick={adicionarGasto}
                className="rounded-2xl bg-blue-600 px-5 py-3 font-medium text-white transition hover:bg-blue-700"
              >
                Adicionar gasto
              </button>

              <button
                onClick={carregarMesAtual}
                className="rounded-2xl bg-slate-700 px-5 py-3 font-medium text-white transition hover:bg-slate-800"
              >
                Recarregar
              </button>

              <button
                onClick={limparMesAtual}
                className="rounded-2xl bg-red-600 px-5 py-3 font-medium text-white transition hover:bg-red-700"
              >
                Limpar mês
              </button>
            </div>

            <div className="mt-8 rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h3 className="text-lg font-semibold text-slate-900">
                    Lista de gastos do mês
                  </h3>
                  <p className="text-sm text-slate-500">
                    Todos os lançamentos da competência atual.
                  </p>
                </div>

                <div className="rounded-full bg-white px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm">
                  {gastos.length} item(ns)
                </div>
              </div>

              {gastos.length === 0 ? (
                <div className="mt-5">
                  <EstadoVazio
                    titulo="Nenhum gasto cadastrado"
                    descricao="Assim que você lançar um gasto, ele aparecerá nesta lista."
                  />
                </div>
              ) : (
                <div className="mt-5 overflow-x-auto rounded-2xl border border-slate-200 bg-white">
                  <table className="min-w-full text-sm">
                    <thead className="bg-slate-100 text-slate-700">
                      <tr>
                        <th className="px-4 py-3 text-left font-semibold">
                          Data
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Hora
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Categoria
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Observação
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Valor
                        </th>
                        <th className="px-4 py-3 text-left font-semibold">
                          Ação
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {gastos.map((gasto, i) => (
                        <tr
                          key={gasto.id ?? `${gasto.data}-${gasto.hora}-${i}`}
                          className="border-t border-slate-200"
                        >
                          <td className="px-4 py-3 font-medium text-slate-800">
                            {formatarDataBR(gasto.data)}
                          </td>
                          <td className="px-4 py-3 text-slate-600">
                            {gasto.hora || "-"}
                          </td>
                          <td className="px-4 py-3">
                            <span
                              className={`rounded-full px-3 py-1 text-xs font-semibold ${classeCategoria(
                                gasto.categoria
                              )}`}
                            >
                              {gasto.categoria}
                            </span>
                          </td>
                          <td className="px-4 py-3 text-slate-600">
                            {gasto.observacao || "Sem observação"}
                          </td>
                          <td className="px-4 py-3 font-semibold text-red-700">
                            {formatarMoeda(gasto.valor)}
                          </td>
                          <td className="px-4 py-3">
                            <button
                              onClick={() => excluirGasto(i)}
                              className="rounded-xl bg-red-600 px-3 py-2 text-white transition hover:bg-red-700"
                            >
                              Excluir
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </section>
        </div>
      </div>
    </main>
  );
}

function CampoBox({
  label,
  children,
}: {
  label: string;
  children: ReactNode;
}) {
  return (
    <div>
      <label className="mb-2 block text-sm font-semibold text-slate-700">
        {label}
      </label>
      {children}
    </div>
  );
}

function HeroMiniCard({
  titulo,
  valor,
}: {
  titulo: string;
  valor: string;
}) {
  return (
    <div className="rounded-2xl border border-white/10 bg-white/10 px-4 py-4 backdrop-blur-sm">
      <p className="text-xs font-medium uppercase tracking-wide text-slate-300">
        {titulo}
      </p>
      <p className="mt-1 text-lg font-bold text-white">{valor}</p>
    </div>
  );
}

function IndicadorResumo({
  titulo,
  valor,
  descricao,
  cor,
}: {
  titulo: string;
  valor: string;
  descricao: string;
  cor: "blue" | "green" | "red" | "violet";
}) {
  const barra =
    {
      blue: "from-blue-600 to-cyan-500",
      green: "from-emerald-600 to-green-500",
      red: "from-rose-600 to-red-500",
      violet: "from-violet-600 to-fuchsia-500",
    }[cor] || "from-slate-700 to-slate-600";

  return (
    <div className="overflow-hidden rounded-[24px] border border-slate-200 bg-white shadow-[0_14px_35px_rgba(15,23,42,0.07)]">
      <div className={`h-2 bg-gradient-to-r ${barra}`} />
      <div className="p-5">
        <p className="text-sm font-semibold text-slate-600">{titulo}</p>
        <p className="mt-2 text-2xl font-bold text-slate-900">{valor}</p>
        <p className="mt-1 text-sm text-slate-500">{descricao}</p>
      </div>
    </div>
  );
}

function MiniResumo({
  titulo,
  valor,
}: {
  titulo: string;
  valor: string;
}) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
      <p className="text-sm text-slate-500">{titulo}</p>
      <p className="mt-1 text-lg font-semibold text-slate-900">{valor}</p>
    </div>
  );
}

function ResumoEscuro({
  titulo,
  valor,
}: {
  titulo: string;
  valor: string;
}) {
  return (
    <div className="rounded-2xl border border-white/10 bg-white/10 p-4">
      <p className="text-sm text-slate-200">{titulo}</p>
      <p className="mt-1 text-lg font-semibold text-white">{valor}</p>
    </div>
  );
}

function EstadoVazio({
  titulo,
  descricao,
}: {
  titulo: string;
  descricao: string;
}) {
  return (
    <div className="rounded-2xl border border-dashed border-slate-300 bg-white px-6 py-8 text-center">
      <p className="text-base font-semibold text-slate-700">{titulo}</p>
      <p className="mt-2 text-sm text-slate-500">{descricao}</p>
    </div>
  );
}

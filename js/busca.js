/* js/busca.js - Lógica do Modal de Pesquisa */

function abrirBusca() {
    const modal = document.getElementById('modalBusca');
    if(modal) {
        modal.style.display = 'block';
        setTimeout(() => document.getElementById('campoBusca').focus(), 100);
    }
}

function fecharBusca() {
    document.getElementById('modalBusca').style.display = 'none';
}

// Fecha ao clicar fora da caixa branca
window.onclick = function(event) {
    const modal = document.getElementById('modalBusca');
    if (event.target == modal) {
        fecharBusca();
    }
}

function realizarBusca() {
    const termo = document.getElementById('campoBusca').value.toLowerCase();
    const divResultados = document.getElementById('resultadosBusca');
    
    // Se digitou pouco, limpa
    if(termo.length < 2) {
        divResultados.innerHTML = '<p class="text-center text-muted mt-3">Digite ano ou edição...</p>';
        return;
    }

    // Verifica se os dados existem
    if(typeof baseDeDados === 'undefined') {
        divResultados.innerHTML = '<p class="text-danger mt-3">Erro: Base de dados não carregada nesta página.</p>';
        return;
    }

    // Filtra (Procura no ANO, na EDIÇÃO ou na DATA)
    const encontrados = baseDeDados.filter(item => {
        return item.ano.toString().includes(termo) || 
               item.edicao.toLowerCase().includes(termo) ||
               item.data.includes(termo);
    });

    // Limita a 20 resultados para não travar
    const topResultados = encontrados.slice(0, 20);

    if(topResultados.length > 0) {
        divResultados.innerHTML = topResultados.map(item => `
            <div class="item-resultado">
                <div>
                    <div class="res-info">Jornal ${item.edicao}</div>
                    <div class="res-data"><i class="bi bi-calendar"></i> ${item.data}</div>
                </div>
                <a href="${item.link}" target="_blank" class="btn-ir">Abrir</a>
            </div>
        `).join('');
        
        if(encontrados.length > 20) {
            divResultados.innerHTML += `<p class="text-center text-muted mt-2 small">...e mais ${encontrados.length - 20} resultados. Refine sua busca.</p>`;
        }
    } else {
        divResultados.innerHTML = '<p class="text-center text-muted mt-3">Nenhum resultado encontrado.</p>';
    }
}
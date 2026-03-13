import json, base64, requests, os, sys
from datetime import datetime


# ── CONFIGURAÇÕES DO GITHUB ──────────────────────────────────
def _ler_token():
    try:
        caminho = os.path.join(os.path.dirname(__file__), "..", "config", "token.txt")
        with open(caminho, "r") as f:
            return f.read().strip()
    except:
        return ""


GITHUB_TOKEN = _ler_token()
GITHUB_USER = "SrMesquitaa"
GITHUB_REPO = "painel-acelera"
GITHUB_FILE = "status.json"


class PainelSync:
    def __init__(self, total: int, data_processada: str):
        self.total = total
        self.data_processada = data_processada
        self.nexus = 0
        self.raio = 0
        self.falhas = 0
        self.leads = []
        self.inicio = _hora()
        self.fim = ""
        self._turno = self._ler_turno()
        self._sha = self._get_sha()

    # ── USO PÚBLICO ─────────────────────────────────────────

    def proximo_turno(self) -> str:
        """Retorna 'nexus' ou 'raio' e já avança o turno interno."""
        equipe = self._turno
        self._turno = "raio" if equipe == "nexus" else "nexus"
        return equipe

    def registrar(self, nome: str, cpf: str, equipe: str, ok: bool):
        """Chame após cada envio."""
        if ok:
            if equipe == "nexus":
                self.nexus += 1
            else:
                self.raio += 1
        else:
            self.falhas += 1
            equipe = "erro"

        self.leads.append({"nome": nome, "cpf": cpf, "equipe": equipe})
        if len(self.leads) > 50:
            self.leads = self.leads[-50:]

    def atualizar(self, status: str = "running", msg: str = ""):
        """Envia o status.json pro GitHub Pages."""
        payload = {
            "status": status,
            "statusText": msg or f"Enviando... {self.nexus + self.raio}/{self.total}",
            "dataProcessada": self.data_processada,
            "inicio": self.inicio,
            "fim": self.fim,
            "total": self.total,
            "nexus": self.nexus,
            "raio": self.raio,
            "falhas": self.falhas,
            "ontemNexus": 449,
            "ontemRaio": 449,
            "ontemMeta": 100,
            "apiWebhook": True,
            "apiSheets": True,
            "apiForms": True,
            "leads": self.leads,
        }

        conteudo = base64.b64encode(
            json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        ).decode("utf-8")

        body = {
            "message": f"painel: {self.data_processada} {_hora()}",
            "content": conteudo,
        }
        if self._sha:
            body["sha"] = self._sha

        try:
            r = requests.put(
                f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/{GITHUB_FILE}",
                json=body,
                headers={"Authorization": f"token {GITHUB_TOKEN}"},
                timeout=10,
            )
            if r.ok:
                self._sha = r.json().get("content", {}).get("sha", self._sha)
        except Exception as e:
            print(f"[painel_sync] Erro ao atualizar GitHub: {e}")

    def concluir(self):
        self.fim = _hora()
        self._salvar_turno(self._turno)
        self.atualizar(
            "stopped",
            f"Concluído — {self.nexus + self.raio} enviados | {self.falhas} falhas",
        )

    # ── INTERNOS ─────────────────────────────────────────────

    def _get_sha(self) -> str:
        try:
            r = requests.get(
                f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/{GITHUB_FILE}",
                headers={"Authorization": f"token {GITHUB_TOKEN}"},
                timeout=10,
            )
            if r.ok:
                return r.json().get("sha", "")
        except:
            pass
        return ""

    @staticmethod
    def _ler_turno() -> str:
        try:
            with open(".ultimo_turno", "r") as f:
                ultimo = f.read().strip()
            return "raio" if ultimo == "nexus" else "nexus"
        except:
            return "nexus"

    @staticmethod
    def _salvar_turno(turno: str):
        try:
            with open(".ultimo_turno", "w") as f:
                f.write(turno)
        except:
            pass


def _hora() -> str:
    return datetime.now().strftime("%H:%M")

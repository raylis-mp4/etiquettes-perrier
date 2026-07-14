import { useState, useMemo } from "react";
import { Droplets, Printer, ChevronRight, Flag, Download } from "lucide-react";
import { QUAND, POTYPES, calculer } from "./rulesEngine.js";
import { telechargerClasseur } from "./excelGenerator.js";

const inputCls =
  "w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 focus:border-teal-600 focus:outline-none focus:ring-2 focus:ring-teal-100";
const labelCls = "mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-500";

export default function App() {
  const [quandId, setQuandId] = useState("premier_po");
  const [premierPoFlavor, setPremierPoFlavor] = useState("nature");
  const [poActuel, setPoActuel] = useState("nature");
  const [poSuivant, setPoSuivant] = useState("nature");
  const [usa, setUsa] = useState(false);
  const [ligne, setLigne] = useState("31");
  const [equipe, setEquipe] = useState("3");
  const [production, setProduction] = useState("FRA/BEL");
  const [article, setArticle] = useState("");
  const [heure, setHeure] = useState("");
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [telechargement, setTelechargement] = useState("idle"); // idle | loading | done

  const { resultats, comboInconnu } = useMemo(
    () => calculer({ quandId, poActuel, poSuivant, premierPoFlavor, usa }),
    [quandId, poActuel, poSuivant, premierPoFlavor, usa]
  );

  const totalEtiquettes = resultats.reduce((sum, r) => sum + r.quantite, 0);

  const dateAffichee = date
    ? new Date(date + "T00:00:00").toLocaleDateString("fr-FR")
    : "……………..";
  const typeApercu = resultats[0]?.etiquette ?? "…………………………";

  const handleTelecharger = async () => {
    setTelechargement("loading");
    try {
      await telechargerClasseur(resultats, { date, heure, ligne, equipe, production, article });
      setTelechargement("done");
      setTimeout(() => setTelechargement("idle"), 1500);
    } catch (e) {
      console.error(e);
      setTelechargement("idle");
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 px-4 py-8 font-sans text-slate-900">
      <div className="mx-auto max-w-3xl">
        <header className="mb-6 flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-teal-700 text-white">
            <Droplets size={20} />
          </div>
          <div>
            <h1 className="text-lg font-bold leading-tight text-slate-900">
              Étiquettes de prélèvement MILIEU L31
            </h1>
          </div>
        </header>

        <div className="grid gap-5 sm:grid-cols-2">
          <section className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-sm font-bold text-slate-700">Situation</h2>

            <div className="space-y-4">
              <div>
                <label className={labelCls}>Quand</label>
                <select className={inputCls} value={quandId} onChange={(e) => setQuandId(e.target.value)}>
                  {QUAND.map((q) => (
                    <option key={q.id} value={q.id}>
                      {q.label}
                    </option>
                  ))}
                </select>
              </div>

              {quandId === "premier_po" && (
                <div>
                  <label className={labelCls}>Nature ou Arôme</label>
                  <select
                    className={inputCls}
                    value={premierPoFlavor}
                    onChange={(e) => setPremierPoFlavor(e.target.value)}
                  >
                    <option value="nature">Nature</option>
                    <option value="arome">Arôme</option>
                  </select>
                </div>
              )}

              {quandId === "changement_article" && (
                <div className="rounded-lg border border-amber-200 bg-amber-50 p-3">
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className={labelCls}>PO actuel</label>
                      <select className={inputCls} value={poActuel} onChange={(e) => setPoActuel(e.target.value)}>
                        {POTYPES.map((t) => (
                          <option key={t.id} value={t.id}>
                            {t.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className={labelCls}>PO suivant</label>
                      <select className={inputCls} value={poSuivant} onChange={(e) => setPoSuivant(e.target.value)}>
                        {POTYPES.map((t) => (
                          <option key={t.id} value={t.id}>
                            {t.label}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                  {comboInconnu && (
                    <p className="mt-2 text-xs font-semibold text-amber-700">
                      Cette combinaison n'existe pas dans le tableau qualité — à vérifier avec le labo.
                    </p>
                  )}
                </div>
              )}

              <label className="flex items-center gap-2 text-sm text-slate-700">
                <Flag size={14} className="text-slate-400" />
                <input type="checkbox" checked={usa} onChange={(e) => setUsa(e.target.checked)} />
                Production USA (+3 étiquettes)
              </label>

              <hr className="border-slate-200" />

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className={labelCls}>Date</label>
                  <input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Heure</label>
                  <input type="time" className={inputCls} value={heure} onChange={(e) => setHeure(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Ligne</label>
                  <input className={inputCls} value={ligne} onChange={(e) => setLigne(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Équipe</label>
                  <input className={inputCls} value={equipe} onChange={(e) => setEquipe(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Production</label>
                  <input className={inputCls} value={production} onChange={(e) => setProduction(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Article</label>
                  <input className={inputCls} value={article} onChange={(e) => setArticle(e.target.value)} />
                </div>
              </div>
            </div>
          </section>

          <section className="flex flex-col gap-5">
            <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="mb-4 flex items-center gap-2 text-sm font-bold text-slate-700">
                <Printer size={16} className="text-teal-700" />
                Étiquettes à imprimer
              </h2>

              {resultats.length === 0 ? (
                <p className="text-sm text-slate-400">Aucune étiquette pour cette combinaison.</p>
              ) : (
                <ul className="space-y-2">
                  {resultats.map((r, i) => (
                    <li key={i} className="flex items-center justify-between rounded-lg bg-slate-50 px-3 py-2 text-sm">
                      <span className="flex items-center gap-2 text-slate-700">
                        <ChevronRight size={14} className="text-teal-600" />
                        {r.etiquette}
                      </span>
                      <span className="rounded-full bg-teal-700 px-2.5 py-0.5 text-xs font-bold text-white">
                        × {r.quantite}
                      </span>
                    </li>
                  ))}
                </ul>
              )}

              <div className="mt-4 flex items-center justify-between border-t border-slate-200 pt-3">
                <span className="text-xs font-semibold uppercase text-slate-500">Total</span>
                <span className="text-lg font-bold text-teal-700">{totalEtiquettes} étiquettes</span>
              </div>

              <button
                onClick={handleTelecharger}
                disabled={totalEtiquettes === 0 || telechargement === "loading"}
                className="mt-4 flex w-full items-center justify-center gap-2 rounded-lg bg-teal-700 px-4 py-2.5 text-sm font-bold text-white transition hover:bg-teal-800 disabled:cursor-not-allowed disabled:bg-slate-300"
              >
                <Download size={16} />
                {telechargement === "loading"
                  ? "Génération..."
                  : telechargement === "done"
                  ? "Téléchargé !"
                  : "Télécharger le fichier Excel"}
              </button>
            </div>

            <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="mb-4 text-sm font-bold text-slate-700">Aperçu d'une étiquette</h2>
              <div className="mx-auto flex w-full max-w-[220px] flex-col items-center justify-center gap-0.5 rounded border border-dashed border-slate-300 bg-slate-50 px-3 py-3 text-center font-sans">
                <p className="text-[9px] font-bold uppercase leading-tight text-slate-800">
                  PRELEVEMENTS
                </p>
                <p className="text-[9px] font-bold uppercase leading-tight text-slate-800">
                  {typeApercu}
                </p>
                <div className="mt-1 w-full space-y-0.5 text-left text-[10px] font-bold leading-tight text-slate-800">
                  <p>Date : {dateAffichee}      Ligne : {ligne || "…"}</p>
                  <p>Heure : {heure || "……..."}      Equipe : {equipe || "…"}</p>
                  <p>Production : {production || "………………….."}</p>
                  <p>Article : {article || "………………….."}</p>
                </div>
              </div>
            </div>
          </section>
        </div>

        <footer className="mt-8 text-center text-[11px] text-slate-400">
          Développé par Youssouf BENLADJAL —{" "}
          <a href="mailto:youssouf.benladjal@gmail.com" className="underline hover:text-slate-500">
            youssouf.benladjal@gmail.com
          </a>
          <br />
          v0.1
        </footer>
      </div>
    </div>
  );
}

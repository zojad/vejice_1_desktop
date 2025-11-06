// src/api/apiVejice.js
import axios from "axios";

/**
 * Pokliče Vejice API in vrne popravljeno poved.
 * Vrnemo samo string (popravljeno besedilo) ali original, če pride do težave.
 */
export async function popraviPoved(poved) {
  const url = "https://gpu-proc1.cjvt.si/popravljalnik-api/popravi_crkovanje";
  const data = {
    vhodna_poved: poved,
    hkratne_napovedi: true,
    ne_označi_imen: false,
    prepričanost_modela: 0.08,
  };
  const config = {
    headers: {
      "Content-Type": "application/json",
      "X-API-KEY": "vejice_API_beta",
    },
    timeout: 15000,
  };

  try {
    const r = await axios.post(url, data, config);
    const d = r?.data || {};
    return (
      d.popravljeno_besedilo ||
      d.popravki?.[0]?.predlog ||
      poved
    );
  } catch (err) {
    // Naj se dodatek ne sesuje – vrni original in zalogiraj.
    console.error("Vejice API error:", err);
    return poved;
  }
}

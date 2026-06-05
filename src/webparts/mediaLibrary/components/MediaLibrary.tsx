import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import type { IMediaLibraryProps } from './IMediaLibraryProps';

// Typ für ein Medienelement aus SharePoint
interface IMediaItem {
  Id: number;
  Title: string;
  Personen: string;
  Tags: string;
  Medientyp: string;
  EigeneTags: string;
  FileRef: string;
  FileLeafRef: string;
}

export default function MediaLibrary({ context }: IMediaLibraryProps) {
  const [items, setItems]       = useState<IMediaItem[]>([]);
  const [suche, setSuche]       = useState('');
  const [loading, setLoading]   = useState(true);
  const [fehler, setFehler]     = useState('');
  const [selected, setSelected] = useState<IMediaItem | null>(null);

  // Daten aus SharePoint laden
  useEffect(() => {
    const sp = spfi().using(SPFx(context));

    sp.web.lists.getByTitle("Medien")
      .items.select("Id", "Title", "Personen", "Tags", "Medientyp", "EigeneTags", "FileRef", "FileLeafRef")()
      .then((data: IMediaItem[]) => {
        setItems(data);
        setLoading(false);
      })
      .catch((err: Error) => {
        setFehler(`Fehler beim Laden: ${err.message}`);
        setLoading(false);
      });
  }, []);

  // Filterlogik: durchsucht Titel, Ort und Tags
  const gefiltert = items.filter(item => {
    if (!suche) return true;
    const q = suche.toLowerCase();
    return (
      item.Title?.toLowerCase().includes(q) ||
      item.Ort?.toLowerCase().includes(q) ||
      item.Tags?.toLowerCase().includes(q) ||
      item.Medientyp?.toLowerCase().includes(q)
    );
  });

  // Hilfsfunktion: Ist die Datei ein Video?
  const istVideo = (dateiname: string): boolean => {
    const ext = dateiname?.split('.').pop()?.toLowerCase() || '';
    return ['mp4', 'mov', 'avi', 'webm'].includes(ext);
  };

  return (
    <div style={styles.container}>

      {/* Kopfzeile */}
      <div style={styles.header}>
        <h2 style={styles.titel}>Medienbibliothek</h2>
        <p style={styles.anzahl}>{gefiltert.length} Einträge</p>
      </div>

      {/* Suchfeld */}
      <input
        type="text"
        placeholder="Nach Name, Ort oder Tag suchen..."
        value={suche}
        onChange={e => setSuche(e.target.value)}
        style={styles.suchfeld}
      />

      {/* Ladeanzeige */}
      {loading && <p style={styles.info}>Wird geladen...</p>}

      {/* Fehlermeldung */}
      {fehler && <p style={styles.fehler}>{fehler}</p>}

      {/* Keine Ergebnisse */}
      {!loading && !fehler && gefiltert.length === 0 && (
        <p style={styles.info}>Keine Medien gefunden.</p>
      )}

      {/* Galerie-Grid */}
      <div style={styles.grid}>
        {gefiltert.map(item => (
          <div
            key={item.Id}
            style={styles.karte}
            onClick={() => setSelected(item)}
          >
            {/* Vorschau: Bild oder Video */}
            {istVideo(item.FileLeafRef) ? (
              <video
                src={item.FileRef}
                style={styles.vorschau}
                muted
              />
            ) : (
              <img
                src={item.FileRef}
                alt={item.Title}
                style={styles.vorschau}
              />
            )}

            {/* Metadaten */}
            <div style={styles.meta}>
  <strong style={styles.name}>{item.Title || item.FileLeafRef}</strong>
  {item.Personen && (
    <span style={styles.tag}>👤 {item.Personen}</span>
  )}
  {item.Medientyp && (
    <span style={styles.tag}>🎞 {item.Medientyp}</span>
  )}
  {item.EigeneTags && item.EigeneTags.split(',').map((t, i) => (
    <span key={i} style={styles.tag}>{t.trim()}</span>
  ))}
</div>
          </div>
        ))}
      </div>

      {/* Lightbox – öffnet sich beim Klick auf ein Medium */}
      {selected && (
        <div style={styles.overlay} onClick={() => setSelected(null)}>
          <div style={styles.lightbox} onClick={e => e.stopPropagation()}>
            <button style={styles.schliessen} onClick={() => setSelected(null)}>✕</button>

            {istVideo(selected.FileLeafRef) ? (
              <video src={selected.FileRef} style={styles.grossMedia} controls autoPlay />
            ) : (
              <img src={selected.FileRef} alt={selected.Title} style={styles.grossMedia} />
            )}

            <div style={styles.lightboxMeta}>
              <h3 style={{ margin: '0 0 8px' }}>{selected.Title || selected.FileLeafRef}</h3>
              {selected.Ort && <p style={{ margin: '4px 0' }}>📍 Ort: {selected.Ort}</p>}
              {selected.Tags && <p style={{ margin: '4px 0' }}>🏷 Tags: {selected.Tags}</p>}
              {selected.Medientyp && <p style={{ margin: '4px 0' }}>Typ: {selected.Medientyp}</p>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// Inline-Styles (kein extra CSS nötig)
const styles: Record<string, React.CSSProperties> = {
  container: {
    fontFamily: 'Segoe UI, sans-serif',
    padding: '16px',
    maxWidth: '1200px',
  },
  header: {
    display: 'flex',
    alignItems: 'baseline',
    gap: '12px',
    marginBottom: '12px',
  },
  titel: {
    margin: 0,
    fontSize: '20px',
    fontWeight: 600,
  },
  anzahl: {
    margin: 0,
    fontSize: '13px',
    color: '#888',
  },
  suchfeld: {
    width: '100%',
    padding: '10px 14px',
    fontSize: '14px',
    border: '1px solid #ddd',
    borderRadius: '8px',
    marginBottom: '20px',
    boxSizing: 'border-box',
    outline: 'none',
  },
  info: {
    color: '#888',
    textAlign: 'center',
    padding: '40px 0',
  },
  fehler: {
    color: '#c00',
    background: '#fff0f0',
    padding: '10px 14px',
    borderRadius: '6px',
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))',
    gap: '16px',
  },
  karte: {
    background: '#fff',
    border: '1px solid #eee',
    borderRadius: '10px',
    overflow: 'hidden',
    cursor: 'pointer',
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
    transition: 'transform 0.15s',
  },
  vorschau: {
    width: '100%',
    height: '150px',
    objectFit: 'cover',
    display: 'block',
    background: '#f5f5f5',
  },
  meta: {
    padding: '10px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  name: {
    fontSize: '13px',
    fontWeight: 600,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  tag: {
    display: 'inline-block',
    background: '#f0f0f0',
    color: '#555',
    fontSize: '11px',
    padding: '2px 8px',
    borderRadius: '20px',
  },
  overlay: {
    position: 'fixed',
    inset: 0,
    background: 'rgba(0,0,0,0.7)',
    zIndex: 1000,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  lightbox: {
    background: '#fff',
    borderRadius: '12px',
    overflow: 'hidden',
    maxWidth: '90vw',
    maxHeight: '90vh',
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
  },
  schliessen: {
    position: 'absolute',
    top: '10px',
    right: '10px',
    background: 'rgba(0,0,0,0.5)',
    color: '#fff',
    border: 'none',
    borderRadius: '50%',
    width: '32px',
    height: '32px',
    fontSize: '16px',
    cursor: 'pointer',
    zIndex: 10,
  },
  grossMedia: {
    maxWidth: '80vw',
    maxHeight: '65vh',
    objectFit: 'contain',
    display: 'block',
  },
  lightboxMeta: {
    padding: '16px',
    fontSize: '14px',
    color: '#333',
  },
};

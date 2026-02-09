package features.shop.model;

public class Riga {
    public String modello, tessuto, trattamento, colore, taglia, genere, tipologia;
    public int giacenza;

    public Riga(String modello, String tessuto, String trattamento, String colore, String taglia, String genere, String tipologia, int giacenza) {
        this.modello = modello;
        this.tessuto = tessuto;
        this.trattamento = trattamento;
        this.colore = colore;
        this.taglia = taglia;
        this.genere = genere;
        this.tipologia = tipologia;
        this.giacenza = giacenza;
    }
}

package features.shop.model;

public class Articolo {
    public String modello, tessuto, trattamento, colore, taglia, genere, tipologia;

    public Articolo(String modello, String tessuto, String trattamento, String colore, String taglia, String genere, String tipologia) {
        this.modello = modello;
        this.tessuto = tessuto;
        this.trattamento = trattamento;
        this.colore = colore;
        this.taglia = taglia;
        this.genere = genere;
        this.tipologia = tipologia;
    }
}

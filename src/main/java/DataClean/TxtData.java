package DataClean;

public class TxtData {
    private String beleg;
    private String referenz;
    private String belegDatum;
    private String bruttoBetrag;
    private String skonto;
    private String zahlBetrag;
    private String sum;

    public String getBeleg() {
        return beleg;
    }

    public TxtData setBeleg(String beleg) {
        this.beleg = beleg;
        return this;
    }

    public String getReferenz() {
        return referenz;
    }

    public TxtData setReferenz(String referenz) {
        this.referenz = referenz;
        return this;
    }

    public String getBelegDatum() {
        return belegDatum;
    }

    public TxtData setBelegDatum(String belegDatum) {
        this.belegDatum = belegDatum;
        return this;
    }

    public String getBruttoBetrag() {
        return bruttoBetrag;
    }

    public TxtData setBruttoBetrag(String bruttoBetrag) {
        this.bruttoBetrag = bruttoBetrag;
        return this;
    }

    public String getSkonto() {
        return skonto;
    }

    public TxtData setSkonto(String skonto) {
        this.skonto = skonto;
        return this;
    }

    public String getZahlBetrag() {
        return zahlBetrag;
    }

    public TxtData setZahlBetrag(String zahlBetrag) {
        this.zahlBetrag = zahlBetrag;
        return this;
    }

    public String getSum() {
        return sum;
    }

    public TxtData setSum(String sum) {
        this.sum = sum;
        return this;
    }

    @Override
    public String toString() {
        return "TxtData{" +
                "beleg='" + beleg + '\'' +
                ", referenz='" + referenz + '\'' +
                ", belegDatum='" + belegDatum + '\'' +
                ", bruttoBetrag='" + bruttoBetrag + '\'' +
                ", skonto='" + skonto + '\'' +
                ", zahlBetrag='" + zahlBetrag + '\'' +
                ", sum='" + sum + '\'' +
                '}';
    }
}

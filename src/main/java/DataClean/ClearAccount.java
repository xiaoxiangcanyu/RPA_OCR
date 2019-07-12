package DataClean;

public class ClearAccount {
    private String id;
    private String account;
    private String reference;
    private String documentNo;
    private String type;
    private String pstngDate;
    private String docDate;
    private String netDueDt;
    private String payT;
    private String curr;
    private String glAmount;
    private String lCurr;
    private String lCamnt;
    private String clrngDoc;
    private String gl;
    private String text;
    private String billDoc;
    private String variance;
    private String varianceMark;
    private String costCenter;
    private String chargeAmount;
    private String paymentDiffAmount;

    public String getId() {
        return id;
    }

    public ClearAccount setId(String id) {
        this.id = id;
        return this;
    }

    public String getAccount() {
        return account;
    }

    public ClearAccount setAccount(String account) {
        this.account = account;
        return this;
    }

    public String getReference() {
        return reference;
    }

    public ClearAccount setReference(String reference) {
        this.reference = reference;
        return this;
    }

    public String getDocumentNo() {
        return documentNo;
    }

    public ClearAccount setDocumentNo(String documentNo) {
        this.documentNo = documentNo;
        return this;
    }

    public String getType() {
        return type;
    }

    public ClearAccount setType(String type) {
        this.type = type;
        return this;
    }

    public String getPstngDate() {
        return pstngDate;
    }

    public ClearAccount setPstngDate(String pstngDate) {
        this.pstngDate = pstngDate;
        return this;
    }

    public String getDocDate() {
        return docDate;
    }

    public ClearAccount setDocDate(String docDate) {
        this.docDate = docDate;
        return this;
    }

    public String getNetDueDt() {
        return netDueDt;
    }

    public ClearAccount setNetDueDt(String netDueDt) {
        this.netDueDt = netDueDt;
        return this;
    }

    public String getPayT() {
        return payT;
    }

    public ClearAccount setPayT(String payT) {
        this.payT = payT;
        return this;
    }

    public String getCurr() {
        return curr;
    }

    public ClearAccount setCurr(String curr) {
        this.curr = curr;
        return this;
    }

    public String getGlAmount() {
        return glAmount;
    }

    public ClearAccount setGlAmount(String glAmount) {
        this.glAmount = glAmount;
        return this;
    }

    public String getlCurr() {
        return lCurr;
    }

    public ClearAccount setlCurr(String lCurr) {
        this.lCurr = lCurr;
        return this;
    }

    public String getlCamnt() {
        return lCamnt;
    }

    public ClearAccount setlCamnt(String lCamnt) {
        this.lCamnt = lCamnt;
        return this;
    }

    public String getClrngDoc() {
        return clrngDoc;
    }

    public ClearAccount setClrngDoc(String clrngDoc) {
        this.clrngDoc = clrngDoc;
        return this;
    }

    public String getGl() {
        return gl;
    }

    public ClearAccount setGl(String gl) {
        this.gl = gl;
        return this;
    }

    public String getText() {
        return text;
    }

    public ClearAccount setText(String text) {
        this.text = text;
        return this;
    }

    public String getBillDoc() {
        return billDoc;
    }

    public ClearAccount setBillDoc(String billDoc) {
        this.billDoc = billDoc;
        return this;
    }

    public String getVariance() {
        return variance;
    }

    public ClearAccount setVariance(String variance) {
        this.variance = variance;
        return this;
    }

    public String getVarianceMark() {
        return varianceMark;
    }

    public ClearAccount setVarianceMark(String varianceMark) {
        this.varianceMark = varianceMark;
        return this;
    }

    public String getCostCenter() {
        return costCenter;
    }

    public ClearAccount setCostCenter(String costCenter) {
        this.costCenter = costCenter;
        return this;
    }

    public String getChargeAmount() {
        return chargeAmount;
    }

    public ClearAccount setChargeAmount(String chargeAmount) {
        this.chargeAmount = chargeAmount;
        return this;
    }

    public String getPaymentDiffAmount() {
        return paymentDiffAmount;
    }

    public ClearAccount setPaymentDiffAmount(String paymentDiffAmount) {
        this.paymentDiffAmount = paymentDiffAmount;
        return this;
    }

    @Override
    public String toString() {
        return "ClearAccount{" +
                "id='" + id + '\'' +
                ", account='" + account + '\'' +
                ", reference='" + reference + '\'' +
                ", documentNo='" + documentNo + '\'' +
                ", type='" + type + '\'' +
                ", pstngDate='" + pstngDate + '\'' +
                ", docDate='" + docDate + '\'' +
                ", netDueDt='" + netDueDt + '\'' +
                ", payT='" + payT + '\'' +
                ", curr='" + curr + '\'' +
                ", glAmount='" + glAmount + '\'' +
                ", lCurr='" + lCurr + '\'' +
                ", lCamnt='" + lCamnt + '\'' +
                ", clrngDoc='" + clrngDoc + '\'' +
                ", gl='" + gl + '\'' +
                ", text='" + text + '\'' +
                ", billDoc='" + billDoc + '\'' +
                ", variance='" + variance + '\'' +
                ", varianceMark='" + varianceMark + '\'' +
                ", costCenter='" + costCenter + '\'' +
                ", chargeAmount='" + chargeAmount + '\'' +
                ", paymentDiffAmount='" + paymentDiffAmount + '\'' +
                '}';
    }
}

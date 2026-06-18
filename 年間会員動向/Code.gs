/**
 * 年間会員動向
 * https://docs.google.com/spreadsheets/d/1eLVE1XRTLq35Rexe_pfSvQ9GRDkguTjbill1LgYR_28/edit
 * ID: 1eLVE1XRTLq35Rexe_pfSvQ9GRDkguTjbill1LgYR_28
 *
 * ブック側の GAS をエクスポートしたらここに上書きする。
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu("年間会員動向").addToUi();
}

(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.PHRASES = api.PHRASES;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  // Дополняйте этот список и нажимайте «Обновить» на chrome://extensions.
  // Сравнение без регистра, буква «ё» считается как «е».
  const PHRASES = [
    "риелторов не беспокоить",
    "риэлторов не беспокоить",
    "риелторам не беспокоить",
    "риэлторам не беспокоить",
    "агентствам не беспокоить",
    "агентам не беспокоить",
    "посредникам не звонить",
    "посредникам не беспокоить",
    "без агентов",
    "без посредников",
    "только собственник",
  ];

  return { PHRASES };
});

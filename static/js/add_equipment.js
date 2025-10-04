(function () {
  const placement = document.getElementById('placement');
  const soldWrap = document.getElementById('sold_to_wrap');
  function toggleSold() {
    if (!placement || !soldWrap) return;
    soldWrap.style.display = (placement.value === 'נמכר') ? '' : 'none';
  }
  if (placement) {
    placement.addEventListener('change', toggleSold);
    toggleSold();
  }
})();
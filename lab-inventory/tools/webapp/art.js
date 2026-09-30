// 吉祥物（自己畫的 SVG，沒有版權問題）：無尾熊「考拉老師」、戴護目鏡的「實驗小鴨」
var ART = {
  // 無尾熊頭像（標誌、問候）
  koala: function (size) {
    return '<svg class="art" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<circle cx="27" cy="42" r="23" fill="#8E9AA8"/><circle cx="93" cy="42" r="23" fill="#8E9AA8"/>' +
      '<circle cx="28" cy="44" r="13" fill="#F6C9D3"/><circle cx="92" cy="44" r="13" fill="#F6C9D3"/>' +
      '<ellipse cx="60" cy="66" rx="40" ry="36" fill="#AEB9C6"/>' +
      '<ellipse cx="60" cy="80" rx="24" ry="17" fill="#C9D2DC"/>' +
      '<circle cx="43" cy="61" r="5" fill="#2B303A"/><circle cx="77" cy="61" r="5" fill="#2B303A"/>' +
      '<circle cx="44.6" cy="59.4" r="1.7" fill="#fff"/><circle cx="78.6" cy="59.4" r="1.7" fill="#fff"/>' +
      '<ellipse cx="60" cy="73" rx="10" ry="12.5" fill="#39404B"/><ellipse cx="56.5" cy="67.5" rx="3" ry="2" fill="#5E6674"/>' +
      '<path d="M54 88q6 4 12 0" stroke="#39404B" stroke-width="2.4" fill="none" stroke-linecap="round"/>' +
      '<ellipse cx="33" cy="76" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/><ellipse cx="87" cy="76" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/>' +
      '</svg>';
  },
  // 睡覺的無尾熊（今天沒有課）
  koalaSleep: function (size) {
    return '<svg class="art" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<circle cx="27" cy="46" r="21" fill="#8E9AA8"/><circle cx="93" cy="46" r="21" fill="#8E9AA8"/>' +
      '<circle cx="28" cy="48" r="12" fill="#F6C9D3"/><circle cx="92" cy="48" r="12" fill="#F6C9D3"/>' +
      '<ellipse cx="60" cy="70" rx="40" ry="34" fill="#AEB9C6"/><ellipse cx="60" cy="84" rx="24" ry="15" fill="#C9D2DC"/>' +
      '<path d="M37 64q6 5 12 0M71 64q6 5 12 0" stroke="#2B303A" stroke-width="2.6" fill="none" stroke-linecap="round"/>' +
      '<ellipse cx="60" cy="76" rx="9.5" ry="11.5" fill="#39404B"/>' +
      '<ellipse cx="33" cy="79" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/><ellipse cx="87" cy="79" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/>' +
      '<text x="92" y="22" font-size="16" font-weight="700" fill="#7C8CA3" font-family="sans-serif">Z</text>' +
      '<text x="104" y="12" font-size="11" font-weight="700" fill="#A3B1C4" font-family="sans-serif">z</text>' +
      '</svg>';
  },
  // 戴護目鏡的小鴨（全部準備好了）
  duck: function (size) {
    return '<svg class="art" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<ellipse cx="60" cy="108" rx="34" ry="5" fill="#000" opacity=".08"/>' +
      '<ellipse cx="50" cy="104" rx="9" ry="4" fill="#FF9B3D"/><ellipse cx="72" cy="104" rx="9" ry="4" fill="#FF9B3D"/>' +
      '<ellipse cx="64" cy="80" rx="36" ry="26" fill="#FFD34E"/>' +
      '<path d="M96 70q12-8 10 8q-4 6-12 4z" fill="#FFD34E"/>' +
      '<ellipse cx="72" cy="80" rx="16" ry="10" fill="#F6BD2A" transform="rotate(-12 72 80)"/>' +
      '<circle cx="50" cy="45" r="24" fill="#FFD34E"/>' +
      '<path d="M50 21q2-9 9-8q-4 3-3 9z" fill="#F6BD2A"/>' +
      '<path d="M24 50q-12 1-13 6q7 5 17 2z" fill="#FF9B3D"/>' +
      '<path d="M36 36q16-6 34 2" stroke="#2F6BE0" stroke-width="4" fill="none" stroke-linecap="round"/>' +
      '<circle cx="47" cy="44" r="9.5" fill="#CFEAFF" stroke="#2F6BE0" stroke-width="3"/>' +
      '<circle cx="47" cy="45" r="3.6" fill="#2B303A"/><circle cx="48.4" cy="43.6" r="1.3" fill="#fff"/>' +
      '<path d="M42 39.5l4-2" stroke="#fff" stroke-width="2" stroke-linecap="round"/>' +
      '<ellipse cx="60" cy="56" rx="5" ry="3" fill="#FF8FA3" opacity=".55"/>' +
      '</svg>';
  },
  // 小鴨拿燒瓶（讀取中）
  duckFlask: function (size) {
    return '<svg class="art bob" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<ellipse cx="58" cy="82" rx="32" ry="24" fill="#FFD34E"/>' +
      '<circle cx="48" cy="48" r="22" fill="#FFD34E"/>' +
      '<path d="M24 52q-11 1-12 6q7 4 16 2z" fill="#FF9B3D"/>' +
      '<circle cx="46" cy="46" r="3.6" fill="#2B303A"/><circle cx="47.3" cy="44.7" r="1.3" fill="#fff"/>' +
      '<ellipse cx="58" cy="58" rx="4.5" ry="2.6" fill="#FF8FA3" opacity=".55"/>' +
      '<path d="M82 52h12M85 52v14l-10 20a5 5 0 0 0 4.5 7h17a5 5 0 0 0 4.5-7l-10-20V52" fill="#E8F6FF" stroke="#2F6BE0" stroke-width="3" stroke-linejoin="round"/>' +
      '<path d="M78.5 80h24.5l3.5 7a3 3 0 0 1-2.7 4.3H77.7A3 3 0 0 1 75 87z" fill="#7DD3A8"/>' +
      '<circle cx="90" cy="44" r="2.5" fill="#7DD3A8"/><circle cx="96" cy="37" r="1.8" fill="#7DD3A8"/>' +
      '<ellipse cx="74" cy="80" rx="12" ry="8" fill="#F6BD2A" transform="rotate(-30 74 80)"/>' +
      '</svg>';
  },
};

(function (window, undefined) {

    "use strict";

    var $ = window.jQuery;
    var document = window.document;

    // SPHostUrl パラメーター名
    var SPHostUrlKey = "SPHostUrl";

    // 現在の URL から SPHostUrl を取得し、ページ内の現在のドメインを指す各リンクに、クエリ文字列として追加します。
    $(document).ready(function () {
        ensureSPHasRedirectedToSharePointRemoved();

        var spHostUrl = getSPHostUrlFromQueryString(window.location.search);
        var currentAuthority = getAuthorityFromUrl(window.location.href).toUpperCase();

        if (spHostUrl && currentAuthority) {
            appendSPHostUrlToLinks(spHostUrl, currentAuthority);
        }
    });

    // 現在のドメインを指すすべてのリンクに、クエリ文字列として SPHostUrl を追加します。
    function appendSPHostUrlToLinks(spHostUrl, currentAuthority) {
        $("a")
            .filter(function () {
                var authority = getAuthorityFromUrl(this.href);
                if (!authority && /^#|:/.test(this.href)) {
                    // サポートされていない他のプロトコルを含むアンカーと URL を、フィルターで除外します。
                    return false;
                }
                return authority.toUpperCase() == currentAuthority;
            })
            .each(function () {
                if (!getSPHostUrlFromQueryString(this.search)) {
                    if (this.search.length > 0) {
                        this.search += "&" + SPHostUrlKey + "=" + spHostUrl;
                    }
                    else {
                        this.search = "?" + SPHostUrlKey + "=" + spHostUrl;
                    }
                }
            });
    }

    // 指定されたクエリ文字列から SPHostUrl を取得します。
    function getSPHostUrlFromQueryString(queryString) {
        if (queryString) {
            if (queryString[0] === "?") {
                queryString = queryString.substring(1);
            }

            var keyValuePairArray = queryString.split("&");

            for (var i = 0; i < keyValuePairArray.length; i++) {
                var currentKeyValuePair = keyValuePairArray[i].split("=");

                if (currentKeyValuePair.length > 1 && currentKeyValuePair[0] == SPHostUrlKey) {
                    return currentKeyValuePair[1];
                }
            }
        }

        return null;
    }

    // 指定された URL が、http/https プロトコルを使用する絶対 URL、またはプロトコルを基準とする相対 URL である場合は、その URL から認証を取得します。
    function getAuthorityFromUrl(url) {
        if (url) {
            var match = /^(?:https:\/\/|http:\/\/|\/\/)([^\/\?#]+)(?:\/|#|$|\?)/i.exec(url);
            if (match) {
                return match[1];
            }
        }
        return null;
    }

    // クエリ文字列内に SPHasRedirectedToSharePoint が存在している場合は、それを削除します。
    // したがって、ユーザーがこの URL をブックマークしている場合は、SPHasRedirectedToSharePoint はその中に含まれません。
    // window.location.search に変更を加えると、サーバーに対する追加の要求が発生することに注意してください。
    function ensureSPHasRedirectedToSharePointRemoved() {
        var SPHasRedirectedToSharePointParam = "&SPHasRedirectedToSharePoint=1";

        var queryString = window.location.search;

        if (queryString.indexOf(SPHasRedirectedToSharePointParam) >= 0) {
            window.location.search = queryString.replace(SPHasRedirectedToSharePointParam, "");
        }
    }

})(window);

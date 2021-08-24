/*
 * http://demon.tw/my-work/javascript-bencode.html
 *
 * Author: Demon
 * Website: http://demon.tw
 * Email: 380401911@qq.com
 */
function decode_int(x, f) {
    f++;
    var newf = x.indexOf('e', f);
    var n = parseInt(x.substring(f,newf));
    if (x.charAt(f) == '-' && x.charAt(f+1) == '0') {
        throw("ValueError");
    } else if (x.charAt(f) == '0' && newf != f+1) {
        throw("ValueError");
    }
    return [n, newf+1];
}

function decode_string(x, f) {
    var colon = x.indexOf(':', f);
    var n = parseInt(x.substring(f,colon));
    if (x.charAt(f) == '0' && colon != f+1) {
        throw("ValueError");
    }
    colon++;
    return [x.substring(colon,colon+n), colon+n];
}

function decode_list(x, f) {
    var r = []; f++;
    while (x.charAt(f) != 'e') {
        var a = decode_func[x.charAt(f)](x, f);
        var v = a[0]; f = a[1];
        r.push(v);
    }
    return [r, f + 1];
}

function decode_dict(x, f) {
    var r = {}; f++;
    while (x.charAt(f) != 'e') {
        var a = decode_string(x, f);
        var k = a[0]; f = a[1];
        a = decode_func[x.charAt(f)](x, f)
        r[k] = a[0]; f = a[1];
    }
    return [r, f + 1];
}

decode_func = {};
decode_func['l'] = decode_list;
decode_func['d'] = decode_dict;
decode_func['i'] = decode_int;
decode_func['0'] = decode_string;
decode_func['1'] = decode_string;
decode_func['2'] = decode_string;
decode_func['3'] = decode_string;
decode_func['4'] = decode_string;
decode_func['5'] = decode_string;
decode_func['6'] = decode_string;
decode_func['7'] = decode_string;
decode_func['8'] = decode_string;
decode_func['9'] = decode_string;

// x is a string containing bencoded data, 
// where each charCodeAt value matches the byte of data
function bdecode(x) {
    try {
        var a = decode_func[x.charAt(0)](x, 0);
        var r = a[0]; var l = a[1];
    } catch(e) {
        throw("not a valid bencoded string");
    }
    if (l != x.length) {
        throw("invalid bencoded value (data after valid prefix)");
    }
    return r;
}



/*
*   
*   JavaScript cannot read binary data directly, 
*   it needs to be processed by the server and returned, 
*   so here is a demonstration of the usage with JScript:
*   
*   function read(path) {
*       var cp1252Chars = [/\u20AC/g,/\u201A/g,/\u0192/g,/\u201E/g,/\u2026/g,/\u2020/g,/\u2021/g,/\u02C6/g,/\u2030/g,/\u0160/g,/\u2039/g,/\u0152/g,/\u017D/g,/\u2018/g,/\u2019/g,/\u201C/g,/\u201D/g,/\u2022/g,/\u2013/g,/\u2014/g,/\u02DC/g,/\u2122/g,/\u0161/g,/\u203A/g,/\u0153/g,/\u017E/g,/\u0178/g];
*       var latin1Chars = ["\u0080","\u0082","\u0083","\u0084","\u0085","\u0086","\u0087","\u0088","\u0089","\u008A","\u008B","\u008C","\u008E","\u0091","\u0092","\u0093","\u0094","\u0095","\u0096","\u0097","\u0098","\u0099","\u009A","\u009B","\u009C","\u009E","\u009F"];
*       var binstream = new ActiveXObject("ADODB.Stream");
*       binstream.Type = 2;
*       binstream.Charset = "iso-8859-1";
*       binstream.Open();
*       binstream.LoadFromFile(path);
*       var s = binstream.ReadText();
*       for (var i = 0; i < 27; i++)
*           s = s.replace(cp1252Chars[i], latin1Chars[i]);
*       return s;
*   }
*   
*   function write(buf, path) {
*       var binstream = new ActiveXObject("ADODB.Stream");
*       binstream.Type = 2;
*       binstream.Charset = "iso-8859-1";
*       binstream.Open();
*       binstream.WriteText(buf);
*       binstream.SaveToFile(path, 2);
*   }
*   
*   var str = read("foo.torrent");
*   try {
*       var dic = bdecode(str);
*   } catch (e) {
*       WScript.Echo(e);
*       WScript.Quit();
*   }
*       /* get the announce url of the tracker
*   var announce = dic["announce"];
*       /* get the name of the torrent
*   var name = dic["info"]["name"];
*       /* get the number of files of the torrent (assuming a multi-file torrent)
*   var number = dic["info"]["files"].length;
    *   /* get the size of the first file of the torrent (assuming a multi-file torrent)
*   var number = dic["info"]["files"][0]["length"];
*       /* change the announce url
*   dic["announce"] = "http://demon.tw";
*       /* and then encode it back to string
*   var new_str = bencode(dic);
*       /* then write it back to a torrent file 
*       /* now the torrent's announce url has been changed to "http://demon.tw"
*    
*   write(new_str, "bar.torrent");
*   
*   
*   /*
*    * Author: Demon
*    * Website: http://demon.tw
*    * Email: 380401911@qq.com
*    *
*
*   function encode_int(x,r) {
*       r.push('i'); r.push(x+''); r.push('e');
*   }
*   
*   function encode_string(x,r) {
*       r.push(x.length+''); r.push(':'); r.push(x);
*   }
*   
*   function encode_list(x,r) {
*       r.push('l');
*       for (var i in x){
*           var type = typeof(x[i]);
*           type = (type == 'object') ? ((x[i] instanceof Array) ? 'list' : 'dict') : type;
*           encode_func[type](x[i], r)
*       }
*       r.push('e');
*   }
*   
*   function encode_dict(x,r) {
*       r.push('d');
*       var keys = [], ilist = {};
*       for (var i in x) {
*           keys.push(i);
*       }
*       keys.sort();
*       for (var j in keys) {
*           ilist[keys[j]] = x[keys[j]];
*       }
*       for (var k in ilist) {
*           r.push(k.length+''); r.push(':'); r.push(k);
*           var v = ilist[k];
*           var type = typeof(v);
*           type = (type == 'object') ? ((v instanceof Array) ? 'list' : 'dict') : type;
*           encode_func[type](v, r);
*       }
*       r.push('e');
*   }
*   
*   encode_func = {};
*   encode_func['number']  = encode_int;
*   encode_func['string']  = encode_string;
*   encode_func['list']    = encode_list;
*   encode_func['dict']    = encode_dict;
*   
*   function bencode(x) {
*       var r = [];
*       var type = typeof(x);
*       type = (type == 'object') ? ((x instanceof Array) ? 'list' : 'dict') : type;
*       encode_func[type](x, r);
*       return r.join('');
*   }
*   
*   
/*
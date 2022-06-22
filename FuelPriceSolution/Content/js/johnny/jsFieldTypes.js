/* Key             Code | Key             Code | Key             Code | Key             Code | Key             Code
   ---------------------|----------------------|----------------------|----------------------|---------------------
   backspace          8 | page up           33 | 0                 48 | a                 65 | k                 75
   tab                9 | page down         34 | 1                 49 | b                 66 | l                 76
   enter             13 | end               35 | 2                 50 | c                 67 | m                 77
   shift             16 | home              36 | 3                 51 | d                 68 | n                 78
   ctrl              17 | left arrow        37 | 4                 52 | e                 69 | o                 79
   alt               18 | up arrow          38 | 5                 53 | f                 70 | p                 80
   pause/brea        19 | right arro        39 | 6                 54 | g                 71 | q                 81
   caps lock         20 | down arrow        40 | 7                 55 | h                 72 | r                 82
   escape            27 | insert            45 | 8                 56 | i                 73 | s                 83
   (space)           32 | delete            46 | 9                 57 | j                 74 | t                 84
   ----------------------------------------------------------------------------------------------------------------
   u                 85 | numpad 1          97 | add              107 | f7               118 | comma            188
   v                 86 | numpad 2          98 | subtract         109 | f8               119 | dash             189
   w                 87 | numpad 3          99 | decimal po       110 | f9               120 | period           190
   x                 88 | numpad 4         100 | divide           111 | f10              121 | forward sl       191
   y                 89 | numpad 5         101 | f1               112 | f11              122 | grave acce       192
   z                 90 | numpad 6         102 | f2               113 | f12              123 | open brack       219
   left windo        91 | numpad 7         103 | f3               114 | num lock         144 | back slash       220
   right wind        92 | numpad 8         104 | f4               115 | scroll loc       145 | close brak       221
   select key        93 | numpad 9         105 | f5               116 | semi-colon       186 | single quo       222
   numpad 0          96 | multiply         106 | f6               117 | equal sign       187 |                       */

function vStandar(event, arr) {
    var _control = event.target.id;
    var _x_ = event.key.toUpperCase();    
    var _chars = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'BACKSPACE', 'ARROWRIGHT', 'ARROWLEFT', 'TAB'];

    if (arr != null) {
        if (_x_ == " ") {
            var _val = $("#" + _control).val() + " ";
            if (_val.length > 1) {
                var _inicio = _val.length - 2;
                var _final = _val.length;
                if (_val.substring(_inicio, _final) == "  ") {
                    event.preventDefault();
                    return;
                }
            }
        }
        if (!_chars.includes(_x_) && !arr.includes(_x_)) {
            event.preventDefault();
        }
    }
    else {
        if (!_chars.includes(_x_)) {
            event.preventDefault();
        }
    }


}

function vAlphabetic(event, arr) {
    var _control = event.target.id;
    var _x_ = event.key.toUpperCase();
    var _chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'BACKSPACE', 'ARROWRIGHT', 'ARROWLEFT', 'TAB'];

    if (arr != null) {
        if (_x_ == " ") {
            var _val = $("#" + _control).val() + " ";
            if (_val.length > 1) {
                var _inicio = _val.length - 2;
                var _final = _val.length;
                if (_val.substring(_inicio, _final) == "  ") {
                    event.preventDefault();
                    return;
                }
            }
        }
        if (!_chars.includes(_x_) && !arr.includes(_x_)) {
            event.preventDefault();
        }
    }
    else {
        if (!_chars.includes(_x_)) {
            event.preventDefault();
        }
    }


}

function vNumeric(event, arr) {
    var _x_ = event.key.toUpperCase();
    var _chars = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', 'BACKSPACE', 'ARROWRIGHT', 'ARROWLEFT', 'TAB'];

    if (arr != null) {
        if (!_chars.includes(_x_) && !arr.includes(_x_)) {
            event.preventDefault();
        }
    }
    else {
        if (!_chars.includes(_x_)) {
            event.preventDefault();
        }
    }


}

function vNokeys(event) {
    var _x_ = event.key.toUpperCase();
    var _chars = ['ARROWRIGHT', 'ARROWLEFT', 'TAB'];
    if (!_chars.includes(_x_)) {
        event.preventDefault();
    }
}

/**
 * DROP:  control.dataTransfer.getData("Text")
 * PASTE: control.clipboardData.getData("Text")
 */

function vDroPas(control, chars, nums, special) {
    var _valor = "";
    if (control.toString() == '[object DragEvent]')
    {
        _valor = control.dataTransfer.getData("Text").toUpperCase();
    }
    else if (control.toString() == '[object ClipboardEvent]')
    {
        _valor = control.clipboardData.getData("Text").toUpperCase();
    }

    if (_valor.includes("  ")) {
        control.preventDefault();
        return;
    }

    var _keys = [];
    var _chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    var _nums = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'];

    if (chars) {
        for (var c = 0; c < _chars.length; c++) {
            _keys.push(_chars[c]);
        }
    }

    if (nums) {
        for (var c = 0; c < _nums.length; c++) {
            _keys.push(_nums[c]);
        }
    }

    if (special != null) {
        for (var c = 0; c < special.length; c++) {
            _keys.push(special[c]);
        }
    }

    for (var c = 0; c < _valor.length; c++) {
        if (!_keys.includes(_valor.charAt(c))) {
            control.preventDefault();
            break;
        }
    }

}
const e = require("cors");

const emptyMap = new Map();
syncAppData = (req, res) => {

    publishAgenda();

    async function publishAgenda() {
        const name = String(req.body.name);
        const type = String(req.body.breaktype);
        if (emptyMap.has(name) && type == emptyMap.get(name)) {
            emptyMap.delete(name);
        }
        else {
            emptyMap.set(name, type);
        }
        const mapString = JSON.stringify(mapToObj(emptyMap));
        res.send(mapString);
    }
    function mapToObj(map) {
        var obj = {}
        map.forEach(function (v, k) {
            obj[k] = v
        })
        return obj
    }
}

module.exports = {
    syncAppData
}
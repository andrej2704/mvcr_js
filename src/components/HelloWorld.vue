<template>
  <div class="hello">
    <input type="text" v-model="search" />
    <button v-on:click="getFile">Search</button>
    <h1>{{ found }}</h1>
    <h1>{{ msg }}</h1>
  </div>
</template>

<script>
import XLSX from "xlsx";
import axios from "axios";
import cheerio from "cheerio";

export default {
  name: "HelloWorld",
  data() {
    return {
      msg: "Search for Visa Status",
      found: "",
      search: ""
    };
  },
  methods: {
    searchInWorkBook: function(workbook) {
      const vm = this;
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      for (var R = range.s.r; R <= range.e.r; ++R) {
        for (var C = range.s.c; C <= range.e.c; ++C) {
          /* find the cell object */
          var cellref = XLSX.utils.encode_cell({ c: C, r: R }); // construct A1 reference for cell
          if (!sheet[cellref]) continue; // if cell doesn't exist, move on
          var cell = sheet[cellref];

          /* if the cell is a text cell with the searched string */
          if (!(cell.t == "s" || cell.t == "str")) continue; // skip if cell is not text
          if (cell.v.includes(vm.search)) {
            vm.found = "Found!!!";
          } else {
            vm.found = "NOT Found!!!";
          }
        }
      }
    },
    getFile: function() {
      let workbook = undefined;
      const vm = this;

      if (!workbook) {
        const prom = axios
          .get("https://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx")
          .then(response => {
            const tmp = response.data;
            const foundurl =
              "https://www.mvcr.cz/" +
              cheerio
                .load(tmp)(".dark")
                .attr("href");
            const url = "https://www.mvcr.cz/soubor/prehled-k-12-11-2018.aspx";
            fetch(foundurl)
              .then(function(res) {
                /* get the data as a Blob */
                if (!res.ok) throw new Error("fetch failed");
                // return res.arrayBuffer();
                return res.arrayBuffer();
              })
              .then(function(ab) {
                /* parse the data when it is received */
                var data = new Uint8Array(ab);
                workbook = XLSX.read(data, { type: "array" });
                vm.searchInWorkBook(workbook);
              });
          });
      } else {
        vm.searchInWorkBook(workbook);
      }
    }
  }
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h1,
h2 {
  font-weight: normal;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>

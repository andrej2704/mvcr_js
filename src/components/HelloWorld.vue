<template>
  <div class="hello">
    <select class="select" v-model="selected">
      <option v-for="option in options" v-bind:value="option.value">
        {{ option.text }}
      </option>
    </select>
    <input type="text" v-model="search" placeholder="OAM-31509" />
    <button v-on:click="mvcrApiSearch">Search</button>
    <div v-if="fileName">Search in {{ fileName }}</div>
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
      search: "",
      workbook: undefined,
      selected: "0",
      options: [
        { text: "DP", value: 0 },
        { text: "ZM", value: 1 },
        { text: "TP", value: 2 }
      ],
      fileName: ""
    };
  },
  methods: {
    mvcrApiSearch: function() {
      const vm = this;
      axios
        .get("https://mvcr-api.azurewebsites.net/", {
          mode: "no-cors",
          params: {
            search: vm.search,
            page: vm.selected
          }
        })
        .then(res => {
          vm.found = res.data.found;
          vm.fileName = res.data.fileName;
        });
    },
    searchInWorkBook: function(workbook) {
      const vm = this;
      const sheet = workbook.Sheets[workbook.SheetNames[vm.selected]];
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
            vm.found = "Found!!! Pick it Up!";
          } else {
            vm.found = "NOT Found!!!";
          }
        }
      }
    },
    getFile: function() {
      const vm = this;

      if (!vm.workbook) {
        axios
          .get("https://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx")
          .then(response => {
            const tmp = response.data;
            const fileName = cheerio
              .load(tmp)(".dark")
              .attr("href");
            vm.fileName = fileName;
            const fullUrl = "https://www.mvcr.cz/" + fileName;
            fetch(fullUrl)
              .then(function(res) {
                /* get the data as a Blob */
                if (!res.ok) throw new Error("fetch failed");
                // return res.arrayBuffer();
                return res.arrayBuffer();
              })
              .then(function(ab) {
                /* parse the data when it is received */
                var data = new Uint8Array(ab);
                vm.workbook = XLSX.read(data, { type: "array" });
                vm.searchInWorkBook(vm.workbook);
              });
          });
      } else {
        vm.searchInWorkBook(vm.workbook);
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

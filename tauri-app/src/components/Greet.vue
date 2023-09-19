<script setup lang="ts">
import {ref} from "vue";
import {invoke} from "@tauri-apps/api/tauri";

import { save } from "@tauri-apps/api/dialog";
import { writeBinaryFile } from "@tauri-apps/api/fs";

const greetMsg = ref("");
const name = ref("");

async function greet() {
    // Learn more about Tauri commands at https://tauri.app/v1/guides/features/command
    greetMsg.value = await invoke("greet", {name: name.value});
}

const write_excel = async () => {
    const result = await invoke("write_excel");
    console.log(result);

    const filename = "download.xlsx";
    const path = await save({ defaultPath: filename });
    if (path) {
      await writeBinaryFile(path, result);
    }

}
</script>

<template>
<form class="row" @submit.prevent="greet">
    <input id="greet-input" v-model="name" placeholder="Enter a name..."/>
    <button type="submit">Greet</button>
</form>

<p>{{ greetMsg }}</p>
<br>
<button @click="write_excel">Write xlsx</button>
</template>

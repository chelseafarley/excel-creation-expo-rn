import { StatusBar } from 'expo-status-bar';
import { StyleSheet, Button, View } from 'react-native';

// expo add expo-file-system expo-sharing xlsx
import * as XLSX from 'xlsx';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';

export default function App() {
  const generateExcel = () => {
    let wb = XLSX.utils.book_new();
    let ws = XLSX.utils.aoa_to_sheet([
      ["Odd", "Even", "Total"],
      [1, 2, { t: 'n', v: 3, f: 'A2+B2'}],
      [3, 4, { t: 'n', v: 7, f: 'A3+B3'}],
      [5, 6, { t: 'n', v: 10, f: 'A4+B4'}]
    ]);

    XLSX.utils.book_append_sheet(wb, ws, "MyFirstSheet", true);

    let ws2 = XLSX.utils.aoa_to_sheet([
      ["Odd*2", "Even*2", "Total"],
      [{t: "n", f: "MyFirstSheet!A2*2"}, {t: "n", f: "MyFirstSheet!B2*2"}, { t: 'n', f: 'A2+B2'}],
      [{t: "n", f: "MyFirstSheet!A3*2"}, {t: "n", f: "MyFirstSheet!B3*2"}, { t: 'n', f: 'A3+B3'}],
      [{t: "n", f: "MyFirstSheet!A4*2"}, {t: "n", f: "MyFirstSheet!B4*2"}, { t: 'n', f: 'A4+B4'}],
    ]);

    XLSX.utils.book_append_sheet(wb, ws2, "MySecondSheet", true);

    const base64 = XLSX.write(wb, { type: "base64" });
    const filename = FileSystem.documentDirectory + "MyExcel.xlsx";
    FileSystem.writeAsStringAsync(filename, base64, {
      encoding: FileSystem.EncodingType.Base64
    }).then(() => {
      Sharing.shareAsync(filename);
    });
  };

  return (
    <View style={styles.container}>
      <Button title="Generate Excel" onPress={generateExcel} />
      <StatusBar style="auto" />
    </View>
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#fff',
    alignItems: 'center',
    justifyContent: 'center',
  },
});

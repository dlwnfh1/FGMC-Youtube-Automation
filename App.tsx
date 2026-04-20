import { SafeAreaView, StatusBar } from "react-native";
import { HomeScreen } from "./src/screens/HomeScreen";
import { palette } from "./src/theme/palette";

export default function App() {
  return (
    <SafeAreaView style={{ flex: 1, backgroundColor: palette.canvas }}>
      <StatusBar barStyle="dark-content" backgroundColor={palette.canvas} />
      <HomeScreen />
    </SafeAreaView>
  );
}


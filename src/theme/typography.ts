import { TextStyle } from "react-native";

type TypeScale = Record<string, TextStyle>;

export const typography: TypeScale = {
  overline: {
    fontSize: 12,
    letterSpacing: 2,
    fontWeight: "700"
  },
  caption: {
    fontSize: 13,
    lineHeight: 18,
    fontWeight: "500"
  },
  body: {
    fontSize: 16,
    lineHeight: 24,
    fontWeight: "400"
  },
  headline: {
    fontSize: 22,
    lineHeight: 28,
    fontWeight: "700"
  },
  title: {
    fontSize: 28,
    lineHeight: 34,
    fontWeight: "700"
  },
  display: {
    fontSize: 40,
    lineHeight: 46,
    fontWeight: "800"
  }
};

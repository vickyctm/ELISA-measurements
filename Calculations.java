import java.lang.*;

/**
 * @author Victoria Torres
 *
 */

public class Calculations extends Inputs {
	// actualizes the factor number
	public static void fix_data(double[] result) {
		double [] d = {50, 150, 450, 1350,
					  4050, 12150, 36540,
					  109350, 328050, 984150};
		for(int i = 0; i < result.length; i++) {
			if(result[i] != 0) {
				factor = (id_dilution / d[i]);
				break;
			}
		}
	}
	
	// removes the sample values that would not result in a negative slope.
	public static double[] fix_negative_slope(double[] data) {
		double p1 = 0;
		double p2 = 0;
		double diff = 0;
		double twenty_percent = 0;
		double[] result = new double [data.length];

		for (int i = 0; i < (data.length-1); i++) {
			p1 = data[i];
			p2 = data[i + 1];
			twenty_percent = (p1 * 0.2);
			diff = (p1 - p2);
			if (diff >= twenty_percent) {
				result[i] = data[i]; //CHECK IF THERE IS AT LEAST 3 NUMBERS IN IT 
			}else result[i] = 0;
		}
		fix_data(result);
		return result;
	}

	// calculates the df
	public static int calculate_df(double[] dilution) {
		return ((int) (dilution[1] / dilution[0]));
	}

	// takes the raw data line and takes the log of it
	public static double[] log_results(double[] data) {
		double[] log = new double[data.length];

		for (int i = 0; i < log.length; i++) {
			if (data[i] > 0) {
				log[i] = Math.log(data[i]);
			} else
				log[i] = -1; // alerts the calculations
		}
		return log;
	}

	public static double Ymean(double[] array) {
		int denominator = 0;
		double sum = 0;

		for (int i = 0; i < array.length; i++) {
			if (array[i] != -1) {
				sum = (array[i] + sum);
				denominator++; // counts the number of numbers log
			}
		}

		return (sum / denominator);
	}

	public static double Xmean(double[] array) {
		double numerator = 0;
		int denominator = 0;

		for (int i = 0; i < array.length; i++) {
			if (array[i] != -1) {
				numerator = (numerator + (i + 1));
				denominator++; // counts the number of numbers log
			}
		}
		return (numerator / denominator);
	}

	public static double sxx(double[] log, double Xmean) {
		double temp = 0;
		double result = 0;

		for (int i = 0; i < log.length; i++) {
			if (log[i] != -1) {
				temp = ((i + 1) - Xmean);
				result = (result + (Math.pow(temp, 2)));
			}
		}
		return result;
	}

	public static double sxy(double[] log, double Xmean, double Ymean) {
		double result = 0;
		for (int i = 0; i < log.length; i++) {
			if (log[i] != -1) {
				result = (result + ((log[i] - Ymean) * ((i + 1) - Xmean)));
			}
		}
		return result;
	}

	public static double slopewPLL(double[] log, double Xmean, double Ymean, double SXX, double SXY) {
		double sxx = sxx(log, Xmean);
		double sxy = sxy(log, Xmean, Ymean);
		return ((SXY + sxy) / (SXX + sxx));
	}

	public static double slope(double[] log, double Xmean, double Ymean) {
		double sxx = sxx(log, Xmean);
		double sxy = sxy(log, Xmean, Ymean);
		return (sxy / sxx);
	}

	public static double wPLL(double rf, double df, double wPLL_slope, double Xmean, double Ymean, double ctrl_Xmean,
			double ctrl_Ymean) {
		double power;
		power = ((Xmean - (Ymean / wPLL_slope)) - (ctrl_Xmean - (ctrl_Ymean / wPLL_slope)));
		return (rf * Math.pow(df, power));
	}

	public static double rfl(double rf, double df, double rfl_denominator, double d_slope, double Xmean, double Ymean) {
		double power;
		power = ((Xmean - (Ymean / d_slope)) - rfl_denominator);
		return (rf * Math.pow(df, power));
	}

	public static double pll(double rf, double df, double pll_slope, double Xmean, double Ymean, double ctrl_Xmean,
			double ctrl_Ymean) {
		double power;
		power = (((Xmean - (Ymean / pll_slope))) - (ctrl_Xmean - (ctrl_Ymean / pll_slope)));
		return (rf * Math.pow(df, power));
	}

	public static double correlation(double[] log) {
		double y_avg = 0;
		double x_avg = 0;
		double x = 0;
		double y = 0;
		double x_den = 0;
		double y_den = 0;
		double numerator = 0;
		double denumerator = 0;

		for (int i = 0; i < log.length; i++) {
			y_avg = (y_avg + log[i]);
			x_avg = (x_avg + (i + 1));
		}
		y_avg = (y_avg / log.length);
		x_avg = (x_avg / log.length);

		for (int i = 0; i < log.length; i++) {
			x = ((i + 1) - x_avg);
			y = (log[i] - y_avg);
			x_den = (x_den + Math.pow(x, 2));
			y_den = (y_den + Math.pow(y, 2));
			numerator = (numerator + (x * y));

		}
		denumerator = (x_den * y_den);
		denumerator = Math.sqrt(denumerator); // same as taking square root

		return Math.pow((numerator / denumerator), 2);
	}

}
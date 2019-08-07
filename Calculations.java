import java.lang.*;

/**
 * @author Victoria Torres
 *
 */

public class Calculations extends Inputs {

	// calculates the df
	public static int calculate_df(double[] dilution) {
		return ((int)(dilution[1] / dilution [0]));
	}

	// takes the raw data line and takes the log of it
	public static double[] log_results(double[] data) {
		double[] log = new double[data.length];

		for (int i = 0; i < log.length; i++) {
			if (data[i] > 0) {
				log[i] = Math.log(data[i]);
			} else log[i] = -1; // alerts the calculations
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
	
	public static double correlation(double[]data) {
		double y_avg = 0;
		double x_avg = 0;
		double x = 0;
		double y = 0;
		double numerator = 0;
		double denumerator = 0;
		
		for(int i = 0; i < data.length; i++) {
			y_avg += data[i]; 
			x_avg += (i+1); 
		}
		y_avg= (y_avg / data.length);
		x_avg= (x_avg / data.length);
		
		for(int i = 0; i < data.length; i++) {
			x +=((i+1) - x_avg);
			y +=(data[i] - y_avg);
			numerator += (((i+1) - x_avg) * (data[i] - y_avg));
			
		}
		denumerator = (x + y);
		denumerator = Math.pow(denumerator, 0.5); //same as taking square root
		
		return Math.pow((numerator/denumerator),2);
	}

}
